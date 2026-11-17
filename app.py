"""MVP skeleton: sites, contacts, protocol import with approval, tasks, outbox."""
import datetime as dt
import hashlib
import secrets
import json
import os
import re
import sqlite3
import tempfile
import uuid
from collections import Counter

import openpyxl
from flask import has_request_context, Flask, abort, flash, g, jsonify, redirect, render_template_string, request, send_from_directory, session, url_for
from markupsafe import Markup, escape
from werkzeug.security import check_password_hash, generate_password_hash

from protocol_parser import extract_rows, extract_supervision_items

DB_PATH = os.environ.get("APP_DB", os.path.join(os.path.dirname(__file__), "app.db"))
SCHEMA = """
CREATE TABLE IF NOT EXISTS sites (
  id INTEGER PRIMARY KEY, name TEXT NOT NULL, address TEXT,
  land_date TEXT, delivery_date TEXT, months_total INTEGER, contract_ref TEXT,
  digest_time TEXT DEFAULT '08:00', weekly_day TEXT DEFAULT 'ראשון'
);
CREATE TABLE IF NOT EXISTS contacts (
  id INTEGER PRIMARY KEY, email TEXT UNIQUE NOT NULL, first_name TEXT, last_name TEXT,
  role TEXT, org TEXT, phone TEXT, mobile TEXT, address TEXT, short_name TEXT,
  pref_channel TEXT DEFAULT 'whatsapp', is_consultant INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS memberships (
  id INTEGER PRIMARY KEY, site_id INTEGER NOT NULL, contact_id INTEGER NOT NULL,
  site_role TEXT NOT NULL, permission TEXT, notify_freq TEXT,
  UNIQUE(site_id, contact_id, site_role)
);
CREATE TABLE IF NOT EXISTS meetings (
  id INTEGER PRIMARY KEY, site_id INTEGER NOT NULL, kind TEXT NOT NULL,
  meeting_date TEXT, source_name TEXT, month_current INTEGER, month_total INTEGER,
  status TEXT DEFAULT 'draft', staged_rows TEXT
);
CREATE TABLE IF NOT EXISTS tasks (
  id INTEGER PRIMARY KEY, site_id INTEGER NOT NULL, public_no TEXT NOT NULL,
  topic_no TEXT, topic TEXT, text TEXT, assignee_text TEXT,
  status TEXT DEFAULT 'פתוחה', target_text TEXT, target_date TEXT,
  last_meeting_id INTEGER, UNIQUE(site_id, public_no)
);
CREATE TABLE IF NOT EXISTS task_targets (
  id INTEGER PRIMARY KEY, task_id INTEGER NOT NULL, target_text TEXT, meeting_id INTEGER
);
CREATE TABLE IF NOT EXISTS task_assignments (
  id INTEGER PRIMARY KEY, task_id INTEGER NOT NULL, contact_id INTEGER NOT NULL,
  raw_name TEXT, confirmed INTEGER DEFAULT 0, UNIQUE(task_id, contact_id)
);
CREATE TABLE IF NOT EXISTS outbox (
  id INTEGER PRIMARY KEY, site_id INTEGER, contact_id INTEGER, channel TEXT,
  body TEXT, mode TEXT DEFAULT 'מיידי', status TEXT DEFAULT 'queued', created TEXT
);
CREATE TABLE IF NOT EXISTS admin_queue (
  id INTEGER PRIMARY KEY, site_id INTEGER, meeting_id INTEGER, item_json TEXT,
  reason TEXT, status TEXT DEFAULT 'פתוח', created TEXT
);
CREATE TABLE IF NOT EXISTS decisions (
  id INTEGER PRIMARY KEY, site_id INTEGER NOT NULL, topic TEXT NOT NULL, text TEXT NOT NULL,
  decided_on TEXT, source TEXT, created TEXT
);
CREATE TABLE IF NOT EXISTS supervision_items (
  id INTEGER PRIMARY KEY, site_id INTEGER NOT NULL, report_date TEXT, item_no TEXT,
  text TEXT, photo TEXT, UNIQUE(site_id, report_date, item_no)
);
CREATE TABLE IF NOT EXISTS task_status_log (
  id INTEGER PRIMARY KEY, task_id INTEGER NOT NULL, status TEXT, handled_by TEXT,
  changed_by TEXT, changed_at TEXT
);
CREATE TABLE IF NOT EXISTS task_closure_pending (
  task_id INTEGER PRIMARY KEY, meeting_id INTEGER, created_at TEXT
);
CREATE TABLE IF NOT EXISTS task_status_options (
  name TEXT PRIMARY KEY, kind TEXT NOT NULL, sort INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS severity_levels (name TEXT PRIMARY KEY, sort INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS item_types (name TEXT PRIMARY KEY, sort INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS task_comments (
  id INTEGER PRIMARY KEY, task_id INTEGER NOT NULL, contact_id INTEGER, text TEXT NOT NULL, created_at TEXT
);
CREATE TABLE IF NOT EXISTS task_images (
  id INTEGER PRIMARY KEY, task_id INTEGER NOT NULL, meeting_id INTEGER, site_id INTEGER,
  file TEXT NOT NULL, topic TEXT, created_at TEXT, UNIQUE(task_id, file)
);
CREATE TABLE IF NOT EXISTS task_docs (
  id INTEGER PRIMARY KEY, task_id INTEGER NOT NULL, name TEXT NOT NULL, created_at TEXT
);
CREATE TABLE IF NOT EXISTS task_links (
  token_hash TEXT PRIMARY KEY, task_id INTEGER NOT NULL, contact_id INTEGER NOT NULL,
  created_at TEXT, expires_at TEXT
);
CREATE TABLE IF NOT EXISTS task_occurrences (
  id INTEGER PRIMARY KEY, task_id INTEGER NOT NULL, meeting_id INTEGER, number TEXT,
  topic TEXT, text TEXT, target TEXT, created_at TEXT
);
CREATE TABLE IF NOT EXISTS plans (
  id INTEGER PRIMARY KEY, site_id INTEGER NOT NULL, title TEXT NOT NULL, version TEXT,
  file_path TEXT NOT NULL, original_name TEXT, uploaded_at TEXT
);
CREATE TABLE IF NOT EXISTS app_settings (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS site_delete_requests (
  token_hash TEXT PRIMARY KEY, site_id INTEGER NOT NULL, requested_by TEXT, created_at TEXT, expires_at TEXT
);
CREATE TABLE IF NOT EXISTS stage_status (
  site_id INTEGER NOT NULL, level TEXT NOT NULL, stage TEXT NOT NULL, status TEXT NOT NULL,
  PRIMARY KEY(site_id, level, stage)
);
"""

app = Flask(__name__)
app.secret_key = os.environ.get("APP_SECRET", "dev-only")


def short_no(value):
    """How a task number is shown: the protocol date only. The stored number is unchanged."""
    return (value or "").split("#")[0]


app.jinja_env.filters["pubno"] = short_no


def db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def _close(_):
    d = g.pop("db", None)
    if d:
        d.close()


TASK_STATUS_SEED = [("פתוחה", "open"), ("בטיפול", "open"), ("ממתינה לאישור", "open"),
                    ("ממתין למפרט", "open"), ("לבדיקה חוזרת", "open"), ("ממתינה לבדיקה", "open"),
                    ("מעוכבת", "open"), ("לידיעה", "info"), ("בוצע", "closed"), ("בוטלה", "closed")]
SEVERITY_SEED = ["קריטי", "גבוה", "בינוני", "נמוך"]
ITEM_TYPE_SEED = ["הערה", "אי-דיוק", "בקשה למפרט", "נושא חזית", "חניית יועץ", "ליקוי ביצוע", "פריט פתוח"]
CLOSED_STATUSES = ("בוצע", "בוטלה")
INFO_STATUSES = ("לידיעה",)


def meeting_day(rows):
    """The protocol's own date (ISO): the most common date among its rows, or None."""
    dates = Counter(parse_date(r.get("meeting_date")) for r in rows if r.get("meeting_date"))
    dates.pop(None, None)
    return dates.most_common(1)[0][0] if dates else None


def split_rows_by_date(rows):
    """A merged file holds several protocols, one per meeting date. Split its rows so each date
    becomes its own staged protocol, oldest first; rows with no readable date stay together last."""
    groups = {}
    for r in rows:
        groups.setdefault(parse_date(r.get("meeting_date")), []).append(r)
    return sorted(groups.items(), key=lambda kv: (kv[0] is None, kv[0] or ""))


def backfill_dates(c):
    """Fill dates that older records never had: meeting dates from their rows, task opening
    dates from the earliest protocol that mentioned them, and publish times from those dates."""
    for mid, staged in c.execute("SELECT id, staged_rows FROM meetings WHERE meeting_date IS NULL").fetchall():
        day = meeting_day(json.loads(staged or "[]"))
        if day:
            c.execute("UPDATE meetings SET meeting_date=? WHERE id=?", (day, mid))
    c.execute("UPDATE meetings SET published_at=meeting_date WHERE published_at IS NULL AND status='published'")
    days = {mid: d for mid, d in c.execute("SELECT id, meeting_date FROM meetings").fetchall()}
    for tid, last in c.execute("SELECT id, last_meeting_id FROM tasks WHERE created_at IS NULL").fetchall():
        ids = {r[0] for r in c.execute("SELECT meeting_id FROM task_targets WHERE task_id=?", (tid,)).fetchall()}
        ids.add(last)
        found = sorted(days[i] for i in ids if days.get(i))
        if found:
            c.execute("UPDATE tasks SET created_at=? WHERE id=?", (found[0], tid))
    c.execute("UPDATE tasks SET published_at=created_at WHERE published_at IS NULL")


def init_db():
    c = sqlite3.connect(DB_PATH)
    c.executescript(SCHEMA)
    c.execute("UPDATE tasks SET status='בוטלה' WHERE status='לא רלוונטי'")
    c.execute("UPDATE task_status_options SET name='בוטלה' WHERE name='לא רלוונטי'")
    for i, (name, kind) in enumerate(TASK_STATUS_SEED):
        c.execute("INSERT OR IGNORE INTO task_status_options(name, kind, sort) VALUES (?,?,?)", (name, kind, i))
    for i, name in enumerate(SEVERITY_SEED):
        c.execute("INSERT OR IGNORE INTO severity_levels(name, sort) VALUES (?,?)", (name, i))
    for i, name in enumerate(ITEM_TYPE_SEED):
        c.execute("INSERT OR IGNORE INTO item_types(name, sort) VALUES (?,?)", (name, i))
    for stmt in ("ALTER TABLE task_assignments ADD COLUMN is_responsible INTEGER DEFAULT 0",
                 "ALTER TABLE tasks ADD COLUMN print_flag INTEGER DEFAULT 0",
                 "ALTER TABLE task_assignments ADD COLUMN is_cc INTEGER DEFAULT 0",
                 "ALTER TABLE memberships ADD COLUMN notify_freq TEXT",
                 "ALTER TABLE sites ADD COLUMN digest_time TEXT DEFAULT '08:00'",
                 "ALTER TABLE sites ADD COLUMN weekly_day TEXT DEFAULT 'ראשון'",
                 "ALTER TABLE outbox ADD COLUMN mode TEXT DEFAULT 'מיידי'",
                 "ALTER TABLE contacts ADD COLUMN password_hash TEXT",
                 "ALTER TABLE contacts ADD COLUMN is_admin INTEGER DEFAULT 0",
                 "ALTER TABLE contacts ADD COLUMN is_guest INTEGER DEFAULT 0",
                 "ALTER TABLE contacts ADD COLUMN guest_token_hash TEXT",
                 "ALTER TABLE contacts ADD COLUMN expires_at TEXT",
                 "ALTER TABLE sites ADD COLUMN levels TEXT",
                 "ALTER TABLE sites ADD COLUMN stages TEXT",
                 "ALTER TABLE tasks ADD COLUMN created_at TEXT",
                 "ALTER TABLE tasks ADD COLUMN severity TEXT",
                 "ALTER TABLE tasks ADD COLUMN item_type TEXT",
                 "ALTER TABLE tasks ADD COLUMN sub_topic TEXT",
                 "ALTER TABLE tasks ADD COLUMN published_at TEXT",
                 "ALTER TABLE meetings ADD COLUMN published_at TEXT",
                 "ALTER TABLE tasks ADD COLUMN approver TEXT",
                 "ALTER TABLE tasks ADD COLUMN closed_at TEXT"):
        try:
            c.execute(stmt)
        except sqlite3.OperationalError:
            pass  # column already exists
    backfill_dates(c)
    if not c.execute("SELECT 1 FROM app_settings WHERE key='all_users_admin'").fetchone():
        c.execute("UPDATE contacts SET is_admin=1 WHERE password_hash IS NOT NULL")
        c.execute("INSERT INTO app_settings(key, value) VALUES ('all_users_admin', ?)",
                  (dt.datetime.now().isoformat(timespec="minutes"),))
    c.commit()
    c.close()


# ---------- dates & priority ----------
_DATE = re.compile(r"^(\d{1,2})[./](\d{1,2})[./](\d{2,4})$")


def parse_date(text):
    m = _DATE.match((text or "").strip())
    if not m:
        return None
    d, mth, y = (int(x) for x in m.groups())
    if y < 100:
        y += 2000
    try:
        return dt.date(y, mth, d).isoformat()
    except ValueError:
        return None


_ANY_DATE = re.compile(r"\d{1,2}[./]\d{1,2}[./]\d{2,4}")


def current_target_date(text):
    """A target field may hold several dates: each protocol that carried the task
    left its estimate here. The last date is the current estimate; the rest is history."""
    dates = [parse_date(d) for d in _ANY_DATE.findall(text or "")]
    dates = [d for d in dates if d]
    return dates[-1] if dates else None


def priority_of(task, today=None):
    """0 = top. Overdue or 'מיידי' tasks come first while still open."""
    if task["status"] == "בוצע":
        return 9
    today = today or dt.date.today().isoformat()
    if task["target_date"] and task["target_date"] < today:
        return 0
    if (task["target_text"] or "").strip() == "מיידי":
        return 1
    return 2


# ---------- importers ----------
CONTACT_EMAIL_HEADER = 'דוא"ל'


def _norm_header(value):
    """Header text with the common Hebrew/typographic quote variants mapped to a plain double quote."""
    text = str(value or "").strip()
    for quote in ("״", "”", "“", "''"):
        text = text.replace(quote, '"')
    return text


def _find_contact_sheet(wb):
    """First worksheet with a row (within the top 10) that contains a 'דוא"ל' column."""
    for ws in wb.worksheets:
        for idx, row in enumerate(ws.iter_rows(values_only=True, max_row=10)):
            header = [_norm_header(h) for h in row]
            if CONTACT_EMAIL_HEADER in header:
                return ws, idx, {h: i for i, h in enumerate(header) if h}
    return None, None, None


ROLE_ALIASES = {"בעל דירה": "בעלי דירות", "בעלי דירה": "בעלי דירות", "בעלים": "בעלי דירות",
                "יזמים": "יזם", "נציג": "נציגות", "קבלן ראשי": "קבלן"}


NOT_A_ROLE = {"רושם", "מכותב"}  # recipient-type words; not site roles


def _site_role_of(value):
    """Site role for a free-text value from the file. Any role text is accepted; the value is
    normalised through ROLE_ALIASES. Returns None for empty text or recipient-type words."""
    text = ROLE_ALIASES.get(value or "", (value or "").strip())
    return None if not text or text in NOT_A_ROLE else text


def role_sections(role, write=False):
    """Sections a site role may use. Roles without a configured rule may view tasks only."""
    table = ROLE_WRITE if write else ROLE_READ
    if role in table:
        return table[role]
    return set() if write else DEFAULT_READ_SECTIONS


def _cell(row, col, name):
    i = col.get(name)
    if i is None or i >= len(row) or row[i] in (None, ""):
        return ""
    return str(row[i]).strip()


EMAIL_PATTERN = r"[^@\s]+@[^@\s]+\.[^@\s]+"
CONTACT_FIELDS = ["email", "first", "last", "role", "org", "phone", "mobile", "address"]


def _contact_rows(path):
    """Every non-empty row of the contacts sheet, normalised, with notes about each row."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws, header_idx, col = _find_contact_sheet(wb)
    if ws is None:
        raise ValueError('לא נמצאה עמודה בשם דוא"ל באף גיליון. ודא שהכותרת מופיעה בשורה הראשונה של הגיליון.')
    out, seen = [], {}
    for offset, r in enumerate(ws.iter_rows(min_row=header_idx + 2, values_only=True)):
        if not any(v not in (None, "") for v in r):
            continue
        email = _cell(r, col, CONTACT_EMAIL_HEADER).lower()
        first, last = _cell(r, col, "שם פרטי"), _cell(r, col, "שם משפחה")
        full = _cell(r, col, "שם מלא")
        if not first and not last and full:
            parts = full.split(None, 1)
            first, last = parts[0], (parts[1] if len(parts) > 1 else "")
        role = _cell(r, col, "תפקיד")
        if not role or role in NOT_A_ROLE:
            role = _cell(r, col, "סוג") if _cell(r, col, "סוג") not in NOT_A_ROLE else role
        line = header_idx + 2 + offset
        notes = []
        if not email:
            notes.append('חסר דוא"ל - לא ייובא')
        elif not re.fullmatch(EMAIL_PATTERN, email):
            notes.append('דוא"ל לא תקין - לא ייובא')
        elif email in seen and seen[email][1] == role:
            notes.append(f"כפילות זהה לשורה {seen[email][0]} (אותו אדם ואותו תפקיד) - מומלץ למחוק אחת מהשורות")
        elif email in seen:
            notes.append(f"מופיע גם בשורה {seen[email][0]} - יתווסף עם התפקיד שבשורה זו")
        if email and email not in seen:
            seen[email] = (line, role)
        if not first and not last:
            notes.append("חסר שם")
        out.append({"line": line, "email": email, "first": first, "last": last, "role": role,
                    "org": _cell(r, col, "ארגון"), "phone": _cell(r, col, "טלפון"),
                    "mobile": _cell(r, col, "טלפון נייד"), "address": _cell(r, col, "כתובת"),
                    "notes": notes})
    return out


def import_rows(rows, site_id=None, site_role=None):
    """Write rows (dicts with CONTACT_FIELDS) as contacts, and as members of the site when a role is known.
    Rows without a usable email are skipped. Returns (added, skipped)."""
    added = skipped = 0
    for row in rows:
        email = (row.get("email") or "").strip().lower()
        if not email or not re.fullmatch(EMAIL_PATTERN, email):
            skipped += 1
            continue
        db().execute(
            """INSERT INTO contacts(email, first_name, last_name, role, org, phone, mobile, address)
               VALUES (?,?,?,?,?,?,?,?)
               ON CONFLICT(email) DO UPDATE SET
                 first_name=COALESCE(NULLIF(contacts.first_name,''), excluded.first_name),
                 last_name=COALESCE(NULLIF(contacts.last_name,''), excluded.last_name),
                 role=COALESCE(NULLIF(contacts.role,''), excluded.role),
                 org=COALESCE(NULLIF(contacts.org,''), excluded.org),
                 phone=COALESCE(NULLIF(contacts.phone,''), excluded.phone),
                 mobile=COALESCE(NULLIF(contacts.mobile,''), excluded.mobile),
                 address=COALESCE(NULLIF(contacts.address,''), excluded.address)""",
            (email, (row.get("first") or "").strip(), (row.get("last") or "").strip(),
             (row.get("role") or "").strip(), (row.get("org") or "").strip(), (row.get("phone") or "").strip(),
             (row.get("mobile") or "").strip(), (row.get("address") or "").strip()))
        # a role cell may list several roles ("יועץ קרקע, מפקח"): each one becomes its own membership
        parts = [_site_role_of(p) for p in (row.get("role") or "").split(",")]
        membership_roles = [r for r in parts if r] or ([site_role] if site_role else [])
        if site_id and membership_roles:
            cid = db().execute("SELECT id FROM contacts WHERE email=?", (email,)).fetchone()[0]
            for membership_role in dict.fromkeys(membership_roles):
                db().execute("INSERT OR IGNORE INTO memberships(site_id, contact_id, site_role) VALUES (?,?,?)",
                             (site_id, cid, membership_role))
        added += 1
    db().commit()
    return added, skipped


def import_contacts(path, site_id=None, site_role=None):
    """Import every usable row of the file. Returns (added, skipped)."""
    return import_rows(_contact_rows(path), site_id=site_id, site_role=site_role)


def _words(text):
    return set(re.findall(r"[\w\u0590-\u05FF]{2,}", (text or "").replace('"', "")))


def same_task(existing, row):
    """Same public number is not enough: numbers restart in every meeting.
    Treat it as the same task only when the topic matches and the wording overlaps."""
    ta, tb = _words(existing["topic"]), _words(row.get("topic"))
    if ta and tb and not (ta & tb):  # topic column is free text; share at least one word
        return False
    a, b = _words(existing["text"]), _words(row.get("task"))
    if len(a) < 4 or len(b) < 4:  # too short to tell apart: never guess
        return False
    return len(a & b) / max(len(a), len(b)) >= 0.4


def to_queue(site_id, meeting_id, row, reason):
    db().execute(
        "INSERT INTO admin_queue(site_id, meeting_id, item_json, reason, created) VALUES (?,?,?,?,?)",
        (site_id, meeting_id, json.dumps(row, ensure_ascii=False), reason,
         dt.datetime.now().isoformat(timespec="minutes")))


def protocol_already_has(task_id, iso_date):
    if not iso_date:
        return False
    return db().execute(
        """SELECT 1 FROM task_occurrences o JOIN meetings mm ON mm.id = o.meeting_id
           WHERE o.task_id=? AND mm.meeting_date=?""", (task_id, iso_date)).fetchone() is not None


def publish_meeting(meeting_id):
    """Publish approved rows. Anything the system cannot match with confidence goes to
    the admin queue instead of being guessed. Rows are processed in meeting-date order."""
    m = db().execute("SELECT * FROM meetings WHERE id=?", (meeting_id,)).fetchone()
    site_id = m["site_id"]
    rows = [r for r in json.loads(m["staged_rows"] or "[]") if not r.get("skip")]

    def key(r):
        d = r.get("meeting_date") or ""
        try:
            return dt.datetime.strptime(d, "%d.%m.%y")
        except ValueError:
            return dt.datetime.min
    rows.sort(key=key)

    published, queued = [], []
    used = set()  # (task, protocol date): one task appears at most once per protocol
    for r in rows:
        number = str(r["number"]).strip()
        if r.get("flags"):
            to_queue(site_id, meeting_id, r, "; ".join(r["flags"]))
            queued.append(number)
            continue
        # Numbers restart in every meeting, so identity comes from content, not from the number.
        day = r.get("meeting_date")
        candidates = [t for t in db().execute("SELECT * FROM tasks WHERE site_id=?", (site_id,)).fetchall()
                      if (t["id"], day) not in used]
        hits = [t for t in candidates if same_task(t, r)]
        if len(hits) > 1:  # a repeated upload of the same protocol date points at the task it already made
            seen = [t for t in hits if protocol_already_has(t["id"], parse_date(day))]
            if len(seen) == 1:
                hits = seen
        if not hits and _words(r.get("task")):
            # Same wording already published from this same protocol date: a repeated upload, not a new item.
            hits = [t for t in candidates if _words(t["text"]) == _words(r.get("task"))
                    and protocol_already_has(t["id"], parse_date(day))]
        if len(hits) > 1:
            to_queue(site_id, meeting_id, r, "התאמה מרובה למשימות קיימות - לאשר לאיזו משימה זה שייך")
            queued.append(number)
            continue
        match = hits[0] if hits else None
        target_date = current_target_date(r["target"])
        if match is None:
            base = f"{r.get('meeting_date') or m['id']}#{number}"
            public = base
            n = 2
            while db().execute("SELECT 1 FROM tasks WHERE site_id=? AND public_no=?", (site_id, public)).fetchone():
                public = f"{base}-{n}"
                n += 1
            cur = db().execute(
                """INSERT INTO tasks(site_id, public_no, topic_no, topic, text, assignee_text,
                       target_text, target_date, last_meeting_id, created_at, severity, item_type, published_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (site_id, public, r["number"], r["topic"], r["task"], r["assignee"],
                 r["target"], target_date, meeting_id,
                 parse_date(r.get("meeting_date")) or dt.date.today().isoformat(),
                 suggest_severity(r["task"]), suggest_item_type(r["task"]),
                 dt.datetime.now().isoformat(timespec="minutes")))
            task_id = cur.lastrowid
        else:
            task_id = match["id"]
            if match["target_text"] != r["target"]:  # new estimate from this protocol
                db().execute(
                    "UPDATE tasks SET target_text=?, target_date=?, last_meeting_id=? WHERE id=?",
                    (r["target"], target_date, meeting_id, task_id))
        done_here = bool(re.search(r"(^|\s)בוצע$", (r.get("target") or "").strip()))  # 'בוצע' or 'dd.mm.yy בוצע'
        if match is not None and match["status"] in CLOSED_STATUSES and not done_here:
            db().execute("UPDATE tasks SET status='לבדיקה חוזרת' WHERE id=?", (task_id,))
            log_task_status(task_id, "לבדיקה חוזרת", "הופיעה שוב בפרוטוקול לאחר סגירה")
        db().execute("INSERT INTO task_targets(task_id, target_text, meeting_id) VALUES (?,?,?)",
                     (task_id, r["target"], meeting_id))
        db().execute(
            "INSERT INTO task_occurrences(task_id, meeting_id, number, topic, text, target, created_at) VALUES (?,?,?,?,?,?,?)",
            (task_id, meeting_id, number, r.get("topic"), r.get("task"), r.get("target"),
             dt.datetime.now().isoformat(timespec="minutes")))
        for name in r.get("images") or []:
            db().execute(
                "INSERT OR IGNORE INTO task_images(task_id, meeting_id, site_id, file, topic, created_at) VALUES (?,?,?,?,?,?)",
                (task_id, meeting_id, site_id, name, (r.get("topic") or "").strip() or "ללא נושא",
                 dt.datetime.now().isoformat(timespec="minutes")))
        if done_here:
            db().execute("UPDATE tasks SET status='בוצע', closed_at=? WHERE id=?", (dt.date.today().isoformat(), task_id))
        published.append(task_id)
        used.add((task_id, day))

    # Decisions written in a protocol ('החלטה') are stored by topic, once per text and date.
    for r in rows:
        if "החלטה" in (r.get("target") or ""):
            topic = r.get("subject") or r.get("topic") or "כללי"
            day = r.get("meeting_date") or ""
            exists = db().execute("SELECT 1 FROM decisions WHERE site_id=? AND text=? AND decided_on=?",
                                  (site_id, r["task"], day)).fetchone()
            if not exists:
                db().execute(
                    "INSERT INTO decisions(site_id, topic, text, decided_on, source, created) VALUES (?,?,?,?,?,?)",
                    (site_id, topic, r["task"], day, "פרוטוקול", dt.datetime.now().isoformat(timespec="minutes")))

    # Done by removal: only for full protocols, only tasks that were in this protocol's
    # listing and are missing from it. Tasks matched into the queue are not closed.
    if m["kind"] in ("ביצוע", "שבועית"):
        listed_ids = set(published)
        # only tasks from the same kind of protocol: a weekly report does not close execution items
        gone = db().execute(
            """SELECT t.id FROM tasks t JOIN meetings lm ON lm.id = t.last_meeting_id
               WHERE t.site_id=? AND t.status NOT IN ('בוצע', 'בוטלה') AND lm.kind=? AND t.last_meeting_id != ?""",
            (site_id, m["kind"], meeting_id)).fetchall()
        for t in gone:
            if t["id"] not in listed_ids:
                db().execute("INSERT OR IGNORE INTO task_closure_pending(task_id, meeting_id, created_at) VALUES (?,?,?)",
                             (t["id"], meeting_id, dt.datetime.now().isoformat(timespec="minutes")))
    for task_id in published:
        db().execute("DELETE FROM task_closure_pending WHERE task_id=?", (task_id,))
        suggest_assignees(task_id)
    db().execute("UPDATE meetings SET status='published', published_at=?, meeting_date=COALESCE(meeting_date, ?) WHERE id=?",
                 (dt.datetime.now().isoformat(timespec="minutes"), meeting_day(rows), meeting_id))
    db().commit()
    return {"published": len(published), "queued": len(queued)}


# ---------- notifications (queued; provider adapter attaches later) ----------
# Membership frequency wins over system defaults:
#   'מיידי' / 'יומי' / 'שבועי' = receive, delivered now / in a daily or weekly digest
#   'כבוי'                     = receive nothing, even if the system would send
#   None                       = system default (supervisors get the site's default schedule)
# Every message without an explicit 'מיידי' goes out at the site's default time
# (sites.digest_time, weekly_day for 'שבועי').
SUPERVISOR_ROLES = ("מפקח", "מנהל מערכת")
FREQS = ("מיידי", "יומי", "שבועי", "כבוי")


def recipients_for_event(site_id):
    rows = db().execute(
        """SELECT m.notify_freq, m.site_role, c.* FROM memberships m
           JOIN contacts c ON c.id = m.contact_id WHERE m.site_id=?""", (site_id,)).fetchall()
    out = []
    for r in rows:
        if r["notify_freq"] == "כבוי":
            continue
        if r["notify_freq"]:
            out.append((r, r["notify_freq"]))
        elif r["site_role"] in SUPERVISOR_ROLES:
            out.append((r, "יומי"))  # site default: goes out at the site's digest time
    return out


# Development phase: every outgoing message (updates, reminders, FYI) goes to the
# developer instead of the real stakeholder. The intended recipient is kept in the text.
DEV_MODE = os.environ.get("APP_DEV_MODE", "1") == "1"
DEV_EMAIL = os.environ.get("APP_DEV_EMAIL", "zeeval@gmail.com")
DEV_MOBILE = os.environ.get("APP_DEV_MOBILE", "054-5700906")  # kept for when WhatsApp is connected


def dev_contact():
    """The one recipient that receives everything while the system is in development:
    the owner's mobile (WhatsApp) when set, otherwise the owner's email."""
    row = db().execute("SELECT * FROM contacts WHERE email=?", (DEV_EMAIL,)).fetchone()
    if row is None:
        db().execute(
            "INSERT INTO contacts(email, first_name, last_name, role, is_consultant, pref_channel, mobile) "
            "VALUES (?,?,?,?,0,'whatsapp',?)", (DEV_EMAIL, "בעל הפרויקט", "", "בדיקות", DEV_MOBILE))
        db().commit()
    elif DEV_MOBILE and row["mobile"] != DEV_MOBILE:
        db().execute("UPDATE contacts SET mobile=?, pref_channel='whatsapp' WHERE id=?",
                     (DEV_MOBILE, row["id"]))
        db().commit()
    return db().execute("SELECT * FROM contacts WHERE email=?", (DEV_EMAIL,)).fetchone()


def enqueue(site_id, contact, body, mode, channel=None):
    if DEV_MODE:
        intended = f"{contact['first_name']} {contact['last_name']}".strip() or contact["email"]
        body = f"[פיתוח - במקור אל: {intended}] {body}"
        contact = dev_contact()
    if DEV_MODE:
        channel = "email"  # until WhatsApp is connected, development messages go to the owner's email
    elif channel is None:
        channel = "email" if contact["is_consultant"] else (contact["pref_channel"] or "whatsapp")
    db().execute(
        "INSERT INTO outbox(site_id, contact_id, channel, body, mode, created) VALUES (?,?,?,?,?,?)",
        (site_id, contact["id"], channel, body, mode, dt.datetime.now().isoformat(timespec="minutes")))


def notify_task_event(task, text):
    for contact, mode in recipients_for_event(task["site_id"]):
        enqueue(task["site_id"], contact, text, mode)
    db().commit()


def notify_supervisors_task_handled(task, by_name, new_status):
    text = (f"משימה {short_no(task['public_no'])} ({task['topic'] or 'ללא נושא'}): {task['text'][:120]} "
            f"- עודכנה לסטטוס '{new_status}' על ידי {by_name}.")
    notify_task_event(task, text)


def send_task_reminders(site_id, today=None):
    """Queue reminders for open tasks that are overdue or due within a day (site default time)."""
    today = today or dt.date.today()
    limit = (today + dt.timedelta(days=1)).isoformat()
    tasks = db().execute(
        """SELECT * FROM tasks WHERE site_id=? AND status NOT IN ('בוצע', 'לידיעה', 'בוטלה')
           AND (target_date IS NOT NULL AND target_date <= ?)""", (site_id, limit)).fetchall()
    for t in tasks:
        when = "באיחור" if t["target_date"] < today.isoformat() else "תאריך היעד הוא מחר או היום"
        text = (f"תזכורת: משימה {short_no(t['public_no'])} ({t['topic'] or 'ללא נושא'}): {t['text'][:120]} - {when}.")
        # reminders go to the site's recipients (by their settings) and to the task's responsible
        # party only; other people on the task (view-only, or copied) keep their link but get no copy
        assignees = db().execute(
            """SELECT c.* FROM task_assignments a JOIN contacts c ON c.id = a.contact_id
               WHERE a.task_id=? AND a.confirmed=1 AND a.is_responsible=1""", (t["id"],)).fetchall()
        seen = set()
        for contact, mode in recipients_for_event(site_id):
            seen.add(contact["id"])
            enqueue(site_id, contact, text, mode)
        for a in assignees:
            if a["id"] not in seen:
                enqueue(site_id, a, text, "יומי")  # site default schedule
                seen.add(a["id"])
    db().commit()


def digest(site_id, mode):
    """Bundle queued messages of one frequency per recipient (run by a daily or weekly job)."""
    rows = db().execute(
        "SELECT contact_id, body FROM outbox WHERE site_id=? AND mode=? AND status='queued' ORDER BY id",
        (site_id, mode)).fetchall()
    bundles = {}
    for r in rows:
        bundles.setdefault(r["contact_id"], []).append(r["body"])
    return {cid: "\n".join(items) for cid, items in bundles.items()}


WEEKDAYS = ("ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת")
_TIME = re.compile(r"^([01]\d|2[0-3]):[0-5]\d$")


@app.route("/sites/<int:site_id>/schedule", methods=["POST"])
def site_schedule(site_id):
    """Site-wide default for scheduled (daily / weekly) messages."""
    t = request.form.get("digest_time", "08:00")
    d = request.form.get("weekly_day", "ראשון")
    if not _TIME.match(t) or d not in WEEKDAYS:
        abort(400)
    db().execute("UPDATE sites SET digest_time=?, weekly_day=? WHERE id=?", (t, d, site_id))
    db().commit()
    return redirect(url_for("site_view", site_id=site_id))


def digest_slots(site):
    """When the scheduled digests of a site go out: daily at digest_time, weekly on weekly_day."""
    return {"יומי": f"כל יום ב-{site['digest_time']}",
            "שבועי": f"כל {site['weekly_day']} ב-{site['digest_time']}"}


@app.route("/memberships/<int:membership_id>/freq", methods=["POST"])
def membership_freq(membership_id):
    freq = request.form.get("freq") or None
    if freq not in (None, *FREQS):
        abort(400)
    db().execute("UPDATE memberships SET notify_freq=? WHERE id=?", (freq, membership_id))
    db().commit()
    return ("", 204)


ADMIN_QUEUE_PAGE = """<!doctype html><html dir="rtl" lang="he"><head><meta charset="utf-8">
<title>תור מנהל מערכת</title>
<style>body{font-family:Arial;background:#fff;color:#222}h1{color:#C62828}
table{border-collapse:collapse;width:100%}th{background:#F2F2F2;color:#6B6B6B}
td,th{border:1px solid #ddd;padding:6px;vertical-align:top}tr.open td{background:#FBE9E9}
button{background:#C62828;color:#fff;border:0;padding:4px 10px}</style></head><body>
<h1>תור מנהל מערכת</h1>
<p>שורות שהמערכת לא זיהתה בוודאות. לא נשמרו כמשימות עד לטיפול.</p>
<table><tr><th>מזהה</th><th>סיבה</th><th>מספר</th><th>נושא</th><th>משימה</th><th>יעד</th><th>פעולה</th></tr>
{% for q in items %}<tr class="open"><td>{{ q['id'] }}</td><td>{{ q['reason'] }}</td><td>{{ q['num'] }}</td>
<td>{{ q['topic'] }}</td><td>{{ q['task'] }}</td><td>{{ q['target'] }}</td>
<td><form method="post" action="{{ url_for('admin_queue_resolve', qid=q['id']) }}">
<button name="action" value="new">הוסף כמשימה חדשה</button> <button name="action" value="drop">סגור בלי יצירה</button>
</form></td></tr>{% else %}<tr><td colspan="7">אין פריטים פתוחים</td></tr>{% endfor %}
</table></body></html>"""


UPLOAD_DIR = os.environ.get("APP_UPLOAD_DIR") or os.path.join(os.path.dirname(__file__), "uploads")
SECTION_NAV = """<p><a href="{{ url_for('site_view', site_id=site_id) }}">חזרה לאתר</a> |
<a href="{{ url_for('decisions_view', site_id=site_id) }}">החלטות</a> |
<a href="{{ url_for('supervision_view', site_id=site_id) }}">דוחות פיקוח</a> |
<a href="{{ url_for('plans_view', site_id=site_id) }}">תוכניות</a> |
<a href="{{ url_for('progress_view', site_id=site_id) }}">התקדמות</a></p>"""
PAGE_CSS = """<meta charset="utf-8"><style>
body{font-family:Arial;color:#222;background:#fff}h1,h2{color:#C62828}
table{border-collapse:collapse;width:100%}th{background:#F2F2F2;color:#6B6B6B}
td,th{border:1px solid #ddd;padding:6px;vertical-align:top}
button{background:#C62828;color:#fff;border:0;padding:6px 12px}
input,textarea,select{padding:4px;margin:2px}
</style>"""


@app.route("/sites/<int:site_id>/decisions", methods=["GET", "POST"])
def decisions_view(site_id):
    if request.method == "POST":
        topic = (request.form.get("topic") or "").strip()
        text = (request.form.get("text") or "").strip()
        if topic and text:
            db().execute(
                "INSERT INTO decisions(site_id, topic, text, decided_on, source, created) VALUES (?,?,?,?,?,?)",
                (site_id, topic, text, (request.form.get("decided_on") or "").strip(),
                 "ידני", dt.datetime.now().isoformat(timespec="minutes")))
            db().commit()
        return redirect(url_for("decisions_view", site_id=site_id))
    rows = db().execute("SELECT * FROM decisions WHERE site_id=? ORDER BY topic, id", (site_id,)).fetchall()
    return render_template_string(
        PAGE_CSS + "<h1>החלטות לפי נושא</h1>" + SECTION_NAV + """
<table><tr><th>נושא</th><th>החלטה</th><th>תאריך</th><th>מקור</th></tr>
{% for r in rows %}<tr><td>{{ r['topic'] }}</td><td>{{ r['text'] }}</td><td>{{ r['decided_on'] }}</td><td>{{ r['source'] }}</td></tr>
{% else %}<tr><td colspan="4">אין החלטות עדיין</td></tr>{% endfor %}</table>
<h2>החלטה חדשה</h2>
<form method="post"><input name="topic" placeholder="נושא (למשל חשמל)" required>
<input name="text" placeholder="ההחלטה" required size="60">
<input name="decided_on" placeholder="תאריך (dd.mm.yy)"><button>שמור</button></form>""",
        rows=rows, site_id=site_id)


@app.route("/sites/<int:site_id>/supervision", methods=["GET", "POST"])
def supervision_view(site_id):
    if request.method == "POST":
        f = request.files.get("file")
        report_date = (request.form.get("report_date") or "").strip()
        if f and report_date:
            os.makedirs(UPLOAD_DIR, exist_ok=True)
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                f.save(tmp.name)
            try:
                for it in extract_supervision_items(tmp.name):
                    db().execute(
                        """INSERT OR REPLACE INTO supervision_items(site_id, report_date, item_no, text, photo)
                           VALUES (?,?,?,?,?)""",
                        (site_id, report_date, it["item_no"], it["text"], it["photo"]))
            finally:
                os.unlink(tmp.name)
            db().commit()
        return redirect(url_for("supervision_view", site_id=site_id))
    rows = db().execute(
        "SELECT * FROM supervision_items WHERE site_id=? ORDER BY report_date, CAST(item_no AS INTEGER)",
        (site_id,)).fetchall()
    return render_template_string(
        PAGE_CSS + "<h1>דוחות פיקוח</h1>" + SECTION_NAV + """
<table><tr><th>תאריך דוח</th><th>פריט</th><th>תיאור</th><th>תמונה</th></tr>
{% for r in rows %}<tr><td>{{ r['report_date'] }}</td><td>{{ r['item_no'] }}</td>
<td>{{ r['text'] }}</td><td>{{ r['photo'] }}</td></tr>
{% else %}<tr><td colspan="4">אין דוחות עדיין</td></tr>{% endfor %}</table>
<h2>העלאת דוח פיקוח</h2>
<form method="post" enctype="multipart/form-data">
<input type="file" name="file" accept="application/pdf" required>
<input name="report_date" placeholder="תאריך הדוח (dd.mm.yy)" required>
<button>העלה</button></form>""",
        rows=rows, site_id=site_id)


@app.route("/sites/<int:site_id>/plans", methods=["GET", "POST"])
def plans_view(site_id):
    if request.method == "POST":
        f = request.files.get("file")
        title = (request.form.get("title") or "").strip()
        if f and title:
            os.makedirs(os.path.join(UPLOAD_DIR, str(site_id)), exist_ok=True)
            ext = os.path.splitext(f.filename or "")[1].lower()
            stored = os.path.join(str(site_id), f"{uuid.uuid4().hex}{ext}")
            f.save(os.path.join(UPLOAD_DIR, stored))
            db().execute(
                """INSERT INTO plans(site_id, title, version, file_path, original_name, uploaded_at)
                   VALUES (?,?,?,?,?,?)""",
                (site_id, title, (request.form.get("version") or "").strip(), stored, f.filename,
                 dt.datetime.now().isoformat(timespec="minutes")))
            db().commit()
        return redirect(url_for("plans_view", site_id=site_id))
    rows = db().execute("SELECT * FROM plans WHERE site_id=? ORDER BY title, id", (site_id,)).fetchall()
    return render_template_string(
        PAGE_CSS + "<h1>תוכניות וגרסאות</h1>" + SECTION_NAV + """
<table><tr><th>שם</th><th>גרסה</th><th>קובץ מקורי</th><th>הועלה</th><th></th></tr>
{% for r in rows %}<tr><td>{{ r['title'] }}</td><td>{{ r['version'] }}</td><td>{{ r['original_name'] }}</td>
<td>{{ r['uploaded_at'] }}</td><td><a href="{{ url_for('plan_file', plan_id=r['id']) }}">הורדה</a></td></tr>
{% else %}<tr><td colspan="5">אין תוכניות עדיין</td></tr>{% endfor %}</table>
<h2>תוכנית חדשה</h2>
<form method="post" enctype="multipart/form-data">
<input name="title" placeholder="שם התוכנית (למשל חתך א-א)" required>
<input name="version" placeholder="גרסה">
<input type="file" name="file" required><button>העלה</button></form>""",
        rows=rows, site_id=site_id)


@app.route("/plans/<int:plan_id>/file")
def plan_file(plan_id):
    r = db().execute("SELECT * FROM plans WHERE id=?", (plan_id,)).fetchone()
    if r is None:
        abort(404)
    return send_from_directory(UPLOAD_DIR, r["file_path"], as_attachment=True,
                               download_name=r["original_name"] or "plan")


# ---------- access: login and roles per site ----------
# Roles on a site (memberships.site_role) and what each may read or change.
# Assumption to confirm with the owner: owners do not see supervision reports.
ROLE_READ = {
    "מפקח": {"tasks", "decisions", "supervision", "plans", "protocols", "progress"},
    "יזם": {"tasks", "decisions", "supervision", "plans", "progress"},
    "נציגות": {"tasks", "decisions", "supervision", "plans", "progress"},
    "בעלי דירות": {"tasks", "decisions", "plans", "progress"},
    "יועצים": {"tasks", "decisions", "plans", "progress"},
    "קבלן": {"tasks", "plans", "progress"},
}
DEFAULT_READ_SECTIONS = {"tasks"}
ROLE_WRITE = {
    "מפקח": {"tasks", "decisions", "supervision", "plans", "protocols", "progress"},
    "יזם": {"decisions"},
    "נציגות": {"decisions"},
    "יועצים": {"plans"},
}

# endpoint -> section it belongs to; endpoints not listed are administrator-only
SECTION_OF = {
    "site_view": "tasks", "site_live": "tasks", "task_status": "tasks", "assignments": "tasks", "confirm_assignee": "tasks",
    "set_responsible": "tasks", "add_responsible": "tasks", "toggle_cc": "tasks", "site_contacts": "tasks", "cc_report": "tasks",
    "site_warnings": "tasks", "site_audit": "tasks", "task_closure": "tasks",
    "task_comment_add": "tasks", "task_doc_add": "tasks", "site_send_links": "tasks", "site_protocol_search": "protocols", "site_protocols": "protocols", "site_protocols_publish": "protocols", "meeting_update": "protocols", "meeting_delete": "protocols", "protocol_image": "protocols", "site_images": "protocols", "assignments_rematch": "tasks", "contact_card": "tasks", "task_triage": "tasks",
    "site_search": "tasks", "task_print_flag": "tasks", "site_print": "tasks", "site_print_clear": "tasks",
    "assign_task": "tasks",
    "decisions_view": "decisions",
    "supervision_view": "supervision",
    "plans_view": "plans", "plan_file": "plans",
    "meeting_upload": "protocols", "meeting_review": "protocols", "meeting_view": "protocols",
    "progress_view": "progress", "progress_save": "progress",
}
# task_link_view is open to anyone holding a valid link; it shows and changes only that one task
PUBLIC_ENDPOINTS = {"login", "logout", "static", "task_link_view", "admin_site_delete_approve", "guest_enter"}
# contact import and editing: administrator and supervisors (deleting field values: administrator only)
MANAGER_ENDPOINTS = {"contacts_import", "contacts_import_preview", "contacts_import_confirm", "contact_edit",
                     "contact_role_add", "contact_role_remove", "contact_delete"}
CONTACT_EDIT_FIELDS = [("first_name", "שם פרטי"), ("last_name", "שם משפחה"), ("org", "ארגון"),
                       ("role", "תפקיד"), ("phone", "טלפון"), ("mobile", "נייד"),
                       ("email", 'דוא"ל'), ("address", "כתובת")]


def is_manager(contact):
    """Administrator, or anyone who is a supervisor (מפקח) on some site."""
    if contact is None:
        return False
    if contact["is_admin"]:
        return True
    return db().execute("SELECT 1 FROM memberships WHERE contact_id=? AND site_role='מפקח'",
                        (contact["id"],)).fetchone() is not None


def current_contact():
    if "cid" not in session:
        return None
    if "_me" not in g:
        g._me = db().execute("SELECT * FROM contacts WHERE id=?", (session["cid"],)).fetchone()
        if g._me is not None and g._me["is_guest"]:
            if not g._me["expires_at"] or g._me["expires_at"] < dt.datetime.now().isoformat(timespec="minutes"):
                session.pop("cid", None)  # the guest window has closed
                g._me = None
    return g._me


def roles_on_site(contact, site_id):
    rows = db().execute("SELECT site_role FROM memberships WHERE site_id=? AND contact_id=?",
                        (site_id, contact["id"])).fetchall()
    return {r["site_role"] for r in rows}


EMAIL_ROLES = {"מפקח", "יזם", "נציגות"}


def can_see_emails(site_id):
    """Participants' email addresses are shown only to administrators and to the supervisor,
    the developer and the committee on this site."""
    me = current_contact()
    if me is None:
        return False
    return bool(me["is_admin"]) or bool(roles_on_site(me, site_id) & EMAIL_ROLES)


LINK_DAYS = 14
LINK_STATUSES_EXCLUDED = ("בוצע", "לידיעה", "בוטלה")  # 'בוצע' is a closure request, see below


def make_task_link(task_id, contact_id):
    """A personal link to one task for one contact. Only the hash is stored."""
    token = secrets.token_urlsafe(24)
    now = dt.datetime.now()
    db().execute("INSERT INTO task_links(token_hash, task_id, contact_id, created_at, expires_at) VALUES (?,?,?,?,?)",
                 (hashlib.sha256(token.encode()).hexdigest(), task_id, contact_id,
                  now.isoformat(timespec="minutes"), (now + dt.timedelta(days=LINK_DAYS)).isoformat(timespec="minutes")))
    return token


def is_confirmed_assignee(contact, task_id):
    if task_id is None:
        return False
    return db().execute("SELECT 1 FROM task_assignments WHERE task_id=? AND contact_id=? AND confirmed=1",
                        (task_id, contact["id"])).fetchone() is not None


def is_task_manager(contact, site_id):
    return bool(contact["is_admin"]) or "מפקח" in roles_on_site(contact, site_id)


def task_status_names(kind=None):
    if kind:
        rows = db().execute("SELECT name FROM task_status_options WHERE kind=? ORDER BY sort", (kind,)).fetchall()
    else:
        rows = db().execute("SELECT name FROM task_status_options ORDER BY sort").fetchall()
    return [r[0] for r in rows]


def suggest_severity(text):
    """Suggested severity from the text only; the supervisor can change it."""
    t = text or ""
    if re.search(r"קריטי|מיידי|בטיחות", t):
        return "קריטי"
    if re.search(r"דחוף|גבוה", t):
        return "גבוה"
    return "בינוני"


def suggest_item_type(text):
    t = text or ""
    if "חזית" in t:
        return "נושא חזית"
    if re.search(r"חני(?:ית|ה)", t):
        return "חניית יועץ"
    if re.search(r"אי[- ]?דיוק|סתירה|אי התאמה", t):
        return "אי-דיוק"
    if "מפרט" in t:
        return "בקשה למפרט"
    if re.search(r"ליקוי|פגם|תיקון", t):
        return "ליקוי ביצוע"
    return "הערה"


def task_repeat_count(task_id):
    """How many different protocols mention this task."""
    return db().execute("SELECT COUNT(DISTINCT meeting_id) FROM task_targets WHERE task_id=?",
                        (task_id,)).fetchone()[0]


def short(text, n=90):
    s = text or ""
    return s if len(s) <= n else s[:n] + "…"


def aged_over_30(task):
    """Open for more than 30 days, by the recorded opening date."""
    try:
        return bool(task["created_at"]) and (dt.date.today() - dt.date.fromisoformat(task["created_at"])).days > 30
    except ValueError:
        return False


def site_kpis(site_id):
    open_names = task_status_names("open")
    closed_names = task_status_names("closed")
    tasks = db().execute("SELECT * FROM tasks WHERE site_id=?", (site_id,)).fetchall()
    today = dt.date.today()
    open_t = [t for t in tasks if t["status"] in open_names]

    def aged(t):
        try:
            return bool(t["created_at"]) and (today - dt.date.fromisoformat(t["created_at"])).days > 30
        except ValueError:
            return False
    reps = {t["id"]: task_repeat_count(t["id"]) for t in open_t}
    return {
        "open": len(open_t),
        "repeats_open": sum(1 for t in open_t if reps[t["id"]] >= 2),
        "critical_open": sum(1 for t in open_t if t["severity"] == "קריטי"),
        "awaiting_spec": sum(1 for t in open_t if t["status"] == "ממתין למפרט"),
        "awaiting_approval": sum(1 for t in open_t if t["status"] == "ממתינה לאישור"),
        "aged_over_30": sum(1 for t in open_t if aged(t)),
        "closed": sum(1 for t in tasks if t["status"] in closed_names),
    }


def site_recommendations(site_id):
    """Action suggestions in the same order the review tool used."""
    open_names = task_status_names("open")
    tasks = db().execute("SELECT * FROM tasks WHERE site_id=?", (site_id,)).fetchall()
    open_t = [t for t in tasks if t["status"] in open_names]
    reps = {t["id"]: task_repeat_count(t["id"]) for t in open_t}
    today = dt.date.today()

    def line(t):
        return short(t["text"])
    recs = []
    for t in sorted([t for t in open_t if reps[t["id"]] >= 2], key=lambda t: -reps[t["id"]])[:6]:
        recs.append(f'לסגור סופית את "{line(t)}" – הופיעה {reps[t["id"]]} פעמים')
    for t in [t for t in open_t if t["severity"] == "קריטי"][:5]:
        recs.append(f"טיפול דחוף בפריט קריטי: {line(t)}")
    for t in [t for t in open_t if t["status"] == "ממתין למפרט"][:5]:
        recs.append(f"להשלים מפרט: {line(t)}")
    for t in [t for t in open_t if t["status"] == "ממתינה לאישור"][:5]:
        recs.append(f"לקבל אישור: {line(t)}")
    aged = []
    for t in open_t:
        try:
            if t["created_at"] and (today - dt.date.fromisoformat(t["created_at"])).days > 30:
                aged.append(t)
        except ValueError:
            pass
    for t in aged[:5]:
        recs.append(f"לעדכן סטטוס או לסגור (פתוח מעל 30 יום): {line(t)}")
    return recs[:20]


def log_task_status(task_id, status, handled_by=""):
    """Record who changed a task's status and when. Outside a request (scheduled work) the
    change is recorded as made by the system."""
    me = current_contact() if has_request_context() else None
    who = (f"{me['first_name'] or ''} {me['last_name'] or ''}".strip() or me["email"]) if me else "מערכת"
    db().execute(
        "INSERT INTO task_status_log(task_id, status, handled_by, changed_by, changed_at) VALUES (?,?,?,?,?)",
        (task_id, status, handled_by, who, dt.datetime.now().isoformat(timespec="minutes")))


@app.before_request
def require_access():
    endpoint = request.endpoint
    if endpoint in PUBLIC_ENDPOINTS or endpoint is None:
        return None
    me = current_contact()
    if me is None:
        return redirect(url_for("login"))
    if me["is_admin"]:
        return None
    if endpoint in MANAGER_ENDPOINTS:
        return None if is_manager(me) else abort(403)
    if endpoint == "task_status" and request.method == "POST" and is_confirmed_assignee(me, (request.view_args or {}).get("task_id")):
        return None
    section = SECTION_OF.get(endpoint)
    if section is None:
        abort(403)  # administrator-only screen
    site_id = (request.view_args or {}).get("site_id")
    if site_id is None:
        # tasks and plan files: find the site from the record
        if "task_id" in (request.view_args or {}):
            row = db().execute("SELECT site_id FROM tasks WHERE id=?", (request.view_args["task_id"],)).fetchone()
            site_id = row["site_id"] if row else None
        elif "plan_id" in (request.view_args or {}):
            row = db().execute("SELECT site_id FROM plans WHERE id=?", (request.view_args["plan_id"],)).fetchone()
            site_id = row["site_id"] if row else None
        elif "meeting_id" in (request.view_args or {}):
            row = db().execute("SELECT site_id FROM meetings WHERE id=?", (request.view_args["meeting_id"],)).fetchone()
            site_id = row["site_id"] if row else None
    if site_id is None:
        abort(404)
    roles = roles_on_site(me, site_id)
    write = request.method == "POST"
    if not any(section in role_sections(r, write) for r in roles):
        abort(403)
    return None


@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""
    if request.method == "POST":
        email = (request.form.get("email") or "").strip().lower()
        pw = request.form.get("password") or ""
        row = db().execute("SELECT * FROM contacts WHERE lower(email)=?", (email,)).fetchone()
        if row and row["password_hash"] and check_password_hash(row["password_hash"], pw):
            session["cid"] = row["id"]
            m = db().execute("SELECT site_id FROM memberships WHERE contact_id=? LIMIT 1", (row["id"],)).fetchone()
            if m is not None:
                return redirect(url_for("site_view", site_id=m["site_id"]))
            return redirect(url_for("sites"))
        error = "פרטי ההתחברות שגויים"
    return render_template_string(
        PAGE_CSS + """<h1>כניסה</h1><p style="color:#C62828">{{ error }}</p>
<form method="post"><input name="email" placeholder="דוא״ל" required>
<input name="password" type="password" placeholder="סיסמה" required>
<button>כניסה</button></form>""", error=error)


@app.route("/logout")
def logout():
    session.pop("cid", None)
    return redirect(url_for("login"))


def set_password_cli(email, password):
    init_db()
    c = sqlite3.connect(DB_PATH)
    cur = c.execute("UPDATE contacts SET password_hash=? WHERE lower(email)=?",
                    (generate_password_hash(password), email.strip().lower()))
    c.commit()
    print("updated" if cur.rowcount else "no such contact")


DEFAULT_STAGES = ["הריסה", "שורינג", "חפירה", "שלד", "מפגשי חזית", "אקוסטיקה",
                  "בניין ירוק", "אינסטלציה", "תשתיות חלשות"]  # from the spec; to be confirmed per site
STATUS_OPTIONS = ("לא התחיל", "בביצוע", "הושלם")
STATUS_BG = {"לא התחיל": "#FFFFFF", "בביצוע": "#FBE9E9", "הושלם": "#F2F2F2"}


def site_levels_stages(site):
    levels = [x.strip() for x in (site["levels"] or "").splitlines() if x.strip()]
    stages = [x.strip() for x in (site["stages"] or "").splitlines() if x.strip()] or DEFAULT_STAGES
    return levels, stages


@app.route("/sites/<int:site_id>/progress", methods=["GET", "POST"])
def progress_view(site_id):
    site = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone()
    if site is None:
        abort(404)
    if request.method == "POST":
        if request.form.get("action") == "config":
            db().execute("UPDATE sites SET levels=?, stages=? WHERE id=?",
                         (request.form.get("levels", ""), request.form.get("stages", ""), site_id))
        else:
            level = request.form.get("level", "")
            stage = request.form.get("stage", "")
            status = request.form.get("status", "")
            if status in STATUS_OPTIONS and level and stage:
                db().execute("INSERT OR REPLACE INTO stage_status(site_id, level, stage, status) VALUES (?,?,?,?)",
                             (site_id, level, stage, status))
        db().commit()
        return redirect(url_for("progress_view", site_id=site_id))
    site = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone()
    levels, stages = site_levels_stages(site)
    cells = {(r["level"], r["stage"]): r["status"] for r in db().execute(
        "SELECT level, stage, status FROM stage_status WHERE site_id=?", (site_id,)).fetchall()}
    writable = request.method == "GET"
    return render_template_string(
        PAGE_CSS + "<h1>התקדמות לפי קומה ושלב</h1>" + SECTION_NAV + """
{% if not levels %}<p>עדיין לא הוגדרו קומות. הגדר אותן למטה, שורה לכל קומה.</p>{% endif %}
<table><tr><th>קומה \\ שלב</th>{% for s in stages %}<th>{{ s }}</th>{% endfor %}</tr>
{% for l in levels %}<tr><th>{{ l }}</th>
{% for s in stages %}{% set st = cells.get((l, s), 'לא התחיל') %}
<td style="background:{{ bg[st] }}"><form method="post">
<input type="hidden" name="level" value="{{ l }}"><input type="hidden" name="stage" value="{{ s }}">
<select name="status" onchange="this.form.submit()">{% for o in options %}
<option {% if o == st %}selected{% endif %}>{{ o }}</option>{% endfor %}</select></form></td>
{% endfor %}</tr>{% endfor %}</table>
<h2>הגדרות האתר</h2>
<form method="post"><input type="hidden" name="action" value="config">
<p>קומות (שורה לכל קומה):</p><textarea name="levels" rows="6" cols="30">{{ site['levels'] or '' }}</textarea>
<p>שלבים (שורה לכל שלב, ריק = ברירת המחדל):</p><textarea name="stages" rows="9" cols="30">{{ site['stages'] or '' }}</textarea>
<p><button>שמור הגדרות</button></p></form>""",
        levels=levels, stages=stages, cells=cells, bg=STATUS_BG, options=STATUS_OPTIONS, site=site,
        site_id=site_id)


@app.route("/admin/queue")
def admin_queue():
    rows = db().execute("SELECT * FROM admin_queue WHERE status='פתוח' ORDER BY id").fetchall()
    items = []
    for r in rows:
        d = json.loads(r["item_json"])
        items.append({"id": r["id"], "reason": r["reason"], "num": d.get("number", ""),
                      "topic": d.get("topic", ""), "task": d.get("task", ""), "target": d.get("target", "")})
    return render_template_string(ADMIN_QUEUE_PAGE, items=items)


@app.route("/admin/queue/<int:qid>/resolve", methods=["POST"])
def admin_queue_resolve(qid):
    """action 'new': the row becomes a new task. action 'drop': the row is left out."""
    q = db().execute("SELECT * FROM admin_queue WHERE id=? AND status='פתוח'", (qid,)).fetchone()
    if q is None:
        abort(404)
    if request.form.get("action") == "new":
        r = json.loads(q["item_json"])
        day = r.get("meeting_date") or ""
        base = f"{day or q['meeting_id']}#{r['number']}"
        public, n = base, 2
        while db().execute("SELECT 1 FROM tasks WHERE site_id=? AND public_no=?", (q["site_id"], public)).fetchone():
            public, n = f"{base}-{n}", n + 1
        db().execute(
            """INSERT INTO tasks(site_id, public_no, topic_no, topic, text, assignee_text,
                   target_text, target_date, last_meeting_id, created_at) VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (q["site_id"], public, r["number"], r["topic"], r["task"], r["assignee"],
             r["target"], current_target_date(r["target"]), q["meeting_id"], dt.date.today().isoformat()))
        new_id = db().execute("SELECT id FROM tasks WHERE site_id=? AND public_no=?",
                              (q["site_id"], public)).fetchone()["id"]
        db().execute("INSERT INTO task_targets(task_id, target_text, meeting_id) VALUES (?,?,?)",
                     (new_id, r["target"], q["meeting_id"]))
        suggest_assignees(new_id)
        status = "נוצרה משימה"
    else:
        status = "נסגר בלי יצירה"
    db().execute("UPDATE admin_queue SET status=? WHERE id=?", (status, qid))
    db().commit()
    return redirect(url_for("admin_queue"))


@app.route("/tasks/<int:task_id>/assign", methods=["POST"])
def assign_task(task_id):
    """Assign the task to a contact chosen from the site's list (the supervisor's decision)."""
    t = db().execute("SELECT site_id FROM tasks WHERE id=?", (task_id,)).fetchone() or abort(404)
    cid = request.form.get("contact_id", type=int)
    member = db().execute("SELECT 1 FROM memberships WHERE site_id=? AND contact_id=?",
                          (t["site_id"], cid)).fetchone() if cid else None
    if not member:
        abort(400)
    db().execute(
        "INSERT OR IGNORE INTO task_assignments(task_id, contact_id, raw_name, confirmed) VALUES (?,?,?,1)",
        (task_id, cid, ""))
    db().execute("UPDATE task_assignments SET confirmed=1 WHERE task_id=? AND contact_id=?", (task_id, cid))
    # the supervisor's choice replaces the other pending suggestions for this task
    db().execute("DELETE FROM task_assignments WHERE task_id=? AND confirmed=0 AND contact_id<>?", (task_id, cid))
    if not db().execute("SELECT 1 FROM task_assignments WHERE task_id=? AND is_responsible=1", (task_id,)).fetchone():
        db().execute("UPDATE task_assignments SET is_responsible=1 WHERE task_id=? AND contact_id=?", (task_id, cid))
    db().commit()
    return _task_reply(task_id, url_for("site_warnings", site_id=t["site_id"]), hide=True,
                       message="המשימה שויכה.")


def _safe_next(url, fallback):
    return url if url and url.startswith("/") and not url.startswith("//") else fallback


@app.route("/tasks/<int:task_id>/print-flag", methods=["POST"])
def task_print_flag(task_id):
    t = db().execute("SELECT site_id FROM tasks WHERE id=?", (task_id,)).fetchone() or abort(404)
    db().execute("UPDATE tasks SET print_flag=CASE WHEN print_flag=1 THEN 0 ELSE 1 END WHERE id=?", (task_id,))
    db().commit()
    return redirect(_safe_next(request.form.get("next"), url_for("site_view", site_id=t["site_id"])))


@app.route("/sites/<int:site_id>/print")
def site_print(site_id):
    """Printable list of the tasks marked for printing."""
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    rows = db().execute("SELECT * FROM tasks WHERE site_id=? AND print_flag=1 ORDER BY public_no",
                        (site_id,)).fetchall()
    return page("להדפסה - " + s["name"], """
<style>@media print { header, .noprint { display: none !important; } main { padding: 0; } }</style>
<p class="noprint"><button onclick="window.print()">הדפסה</button>
<form method="post" action="{{ url_for('site_print_clear', site_id=s['id']) }}" style="display:inline">
<button onclick="return confirm('לבטל את כל הסימונים להדפסה?')">ביטול כל הסימונים</button></form></p>
<h3>משימות להדפסה ({{ rows|length }})</h3>
<table><tr><th>#</th><th>נושא</th><th>משימה</th><th>לטיפול (כפי שכתוב)</th><th>יעד</th><th>סטטוס</th></tr>
{% for t in rows %}<tr><td>{{ t['public_no']|pubno }}</td><td>{{ t['topic'] or '' }}</td><td>{{ t['text'] }}</td>
<td>{{ t['assignee_text'] or '' }}</td><td>{{ t['target_text'] or '' }}</td><td>{{ t['status'] }}</td></tr>
{% else %}<tr><td colspan="6">לא סומנו משימות להדפסה</td></tr>{% endfor %}</table>
""", rows=rows, s=s, url_for=url_for)


@app.route("/sites/<int:site_id>/print/clear", methods=["POST"])
def site_print_clear(site_id):
    db().execute("UPDATE tasks SET print_flag=0 WHERE site_id=?", (site_id,))
    db().commit()
    return redirect(url_for("site_print", site_id=site_id))


def _wants_json():
    """Forms sent by the page's script (fetch) want a result, not a redirect."""
    return request.headers.get("X-Requested-With") == "fetch"


def _task_reply(task_id, redirect_to, message="", hide=False):
    """Browser forms keep the usual redirect. Script forms get the new state so the row updates in place."""
    if not _wants_json():
        return redirect(redirect_to)
    row = db().execute("SELECT status FROM tasks WHERE id=?", (task_id,)).fetchone()
    st = row["status"] if row else None
    hide = hide or st is None or st in CLOSED_STATUSES + INFO_STATUSES
    return jsonify(ok=True, status=st, css=STATUS_CLASS.get(st, ""), hide=hide, message=message)


@app.route("/tasks/<int:task_id>/triage", methods=["POST"])
def task_triage(task_id):
    """For a task with no one to assign: 'info' marks it for information only, 'delete' removes it."""
    t = db().execute("SELECT site_id FROM tasks WHERE id=?", (task_id,)).fetchone() or abort(404)
    action = request.form.get("action")
    if action == "info":
        db().execute("UPDATE tasks SET status='לידיעה' WHERE id=?", (task_id,))
        log_task_status(task_id, "לידיעה")
    elif action == "delete":
        db().execute("DELETE FROM task_assignments WHERE task_id=?", (task_id,))
        db().execute("DELETE FROM task_targets WHERE task_id=?", (task_id,))
        db().execute("DELETE FROM tasks WHERE id=?", (task_id,))
    else:
        abort(400)
    db().commit()
    return _task_reply(task_id, url_for("site_warnings", site_id=t["site_id"]), hide=True,
                       message="המשימה נמחקה." if action == "delete" else "המשימה סומנה לידיעה.")


@app.route("/tasks/<int:task_id>/status", methods=["POST"])
def task_status(task_id):
    t = db().execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone() or abort(404)
    new_status = request.form["status"]
    if new_status not in task_status_names():
        abort(400)
    me = current_contact()
    if not is_task_manager(me, t["site_id"]):
        # a confirmed assignee reports handling; closing is only a request for the supervisor
        if not is_confirmed_assignee(me, task_id):
            abort(403)
        if new_status == "בוצע":
            db().execute("INSERT OR IGNORE INTO task_closure_pending(task_id, meeting_id, created_at) VALUES (?,?,?)",
                         (task_id, t["last_meeting_id"], dt.datetime.now().isoformat(timespec="minutes")))
            log_task_status(task_id, "בוצע", "בקשת סגירה ממתינה לאישור המפקח")
            db().commit()
            return _task_reply(task_id, url_for("site_view", site_id=t["site_id"]),
                               message="נשלחה בקשת סגירה למפקח לאישור.")
        if new_status in LINK_STATUSES_EXCLUDED:
            abort(400)
    by_name = request.form.get("by", "לא צוין")
    if new_status in CLOSED_STATUSES:
        db().execute("UPDATE tasks SET status=?, closed_at=? WHERE id=?",
                     (new_status, dt.date.today().isoformat(), task_id))
    else:
        db().execute("UPDATE tasks SET status=? WHERE id=?", (new_status, task_id))
    log_task_status(task_id, new_status, request.form.get("by", ""))
    db().commit()
    if new_status != "פתוחה":
        notify_supervisors_task_handled(t, by_name, new_status)
    return _task_reply(task_id, url_for("site_view", site_id=t["site_id"]), message="הסטטוס עודכן.")


# ---------- assignee linking: text in 'לטיפול' -> contacts (supervisor confirms) ----------
def _full_names(site_id):
    rows = db().execute(
        """SELECT DISTINCT c.id, c.first_name, c.last_name FROM memberships m
           JOIN contacts c ON c.id = m.contact_id WHERE m.site_id=?""", (site_id,)).fetchall()
    return rows


def rematch_site(site_id):
    """Run name matching again for the site's open tasks (after contacts were imported or changed)."""
    for t in db().execute("SELECT id FROM tasks WHERE site_id=? AND status NOT IN ('בוצע', 'בוטלה')", (site_id,)).fetchall():
        suggest_assignees(t["id"])
    db().commit()


def suggest_assignees(task_id):
    """Find contacts of the site whose name appears in the task's 'לטיפול' text.
    A full-name match (first and last name) is linked at once. A partial match ('ליאור י', first
    name only) is linked at once only when exactly one person on the site has that first name;
    otherwise it waits for the supervisor. The first linked person becomes responsible."""
    t = db().execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
    text = t["assignee_text"] or ""
    found = []
    people = list(_full_names(t["site_id"]))
    firsts, initials = {}, {}
    for c in people:
        key = (c["first_name"] or "").strip()
        firsts[key] = firsts.get(key, 0) + 1
        ikey = (key, (c["last_name"] or "").strip()[:1])
        initials[ikey] = initials.get(ikey, 0) + 1
    hits = []  # (contact_id, matched text, first name, kind) with kind: full / initial / first
    for c in people:
        first, last = (c["first_name"] or "").strip(), (c["last_name"] or "").strip()
        if not first:
            continue
        candidates = []
        if last:
            candidates.append((f"{first} {last}", "full"))
            candidates.append((f"{first} {last[:1]}", "initial"))
        candidates.append((first, "first"))
        for cand, kind in candidates:
            if re.search(rf"(?<!\w){re.escape(cand)}(?!\w)", text):
                hits.append((c["id"], cand, first, kind))
                break
    # "ליאור י" fits a person with that initial; a bare "ליאור" must not also be suggested
    specific_firsts = {first for _, _, first, kind in hits if kind != "first"}
    for cid, cand, first, kind in hits:
        if kind == "first" and first in specific_firsts:
            continue
        contact = next(c for c in people if c["id"] == cid)
        ikey = (first, (contact["last_name"] or "").strip()[:1])
        # "ליאור י" is the same person as "ליאור ישראלי" when no one else on the site has that initial
        auto = kind == "full" or (kind == "first" and firsts.get(first, 0) == 1) \
            or (kind == "initial" and initials.get(ikey, 0) == 1)
        found.append((cid, cand, auto))
    # once someone is confirmed on the task, only confirmed matches are added again (no new suggestions)
    already_confirmed = db().execute(
        "SELECT 1 FROM task_assignments WHERE task_id=? AND confirmed=1", (task_id,)).fetchone() is not None
    for cid, raw, full in found:
        if already_confirmed and not full:
            continue
        db().execute(
            "INSERT OR IGNORE INTO task_assignments(task_id, contact_id, raw_name, confirmed) VALUES (?,?,?,?)",
            (task_id, cid, raw, 1 if full else 0))
        if full:
            db().execute("UPDATE task_assignments SET confirmed=1 WHERE task_id=? AND contact_id=?", (task_id, cid))
            has_responsible = db().execute(
                "SELECT 1 FROM task_assignments WHERE task_id=? AND is_responsible=1", (task_id,)).fetchone()
            if not has_responsible:
                db().execute("UPDATE task_assignments SET is_responsible=1 WHERE task_id=? AND contact_id=?",
                             (task_id, cid))
    db().commit()
    return found


@app.route("/tasks/<int:task_id>/assignees/<int:contact_id>/confirm", methods=["POST"])
def confirm_assignee(task_id, contact_id):
    db().execute("UPDATE task_assignments SET confirmed=1 WHERE task_id=? AND contact_id=?",
                 (task_id, contact_id))
    # the first confirmed person on a task becomes its responsible party
    has_responsible = db().execute(
        "SELECT 1 FROM task_assignments WHERE task_id=? AND is_responsible=1", (task_id,)).fetchone()
    if not has_responsible:
        db().execute("UPDATE task_assignments SET is_responsible=1 WHERE task_id=? AND contact_id=?",
                     (task_id, contact_id))
    db().commit()
    t = db().execute("SELECT site_id FROM tasks WHERE id=?", (task_id,)).fetchone()
    return redirect(url_for("assignments", site_id=t["site_id"]))


@app.route("/tasks/<int:task_id>/assignees/<int:contact_id>/responsible", methods=["POST"])
def set_responsible(task_id, contact_id):
    """Make one confirmed person the responsible party; everyone else on the task is view-only."""
    t = db().execute("SELECT site_id FROM tasks WHERE id=?", (task_id,)).fetchone() or abort(404)
    confirmed = db().execute(
        "SELECT 1 FROM task_assignments WHERE task_id=? AND contact_id=? AND confirmed=1",
        (task_id, contact_id)).fetchone()
    if not confirmed:
        abort(400)  # only a confirmed assignee can be responsible
    db().execute("UPDATE task_assignments SET is_responsible=0 WHERE task_id=?", (task_id,))
    db().execute("UPDATE task_assignments SET is_responsible=1 WHERE task_id=? AND contact_id=?",
                 (task_id, contact_id))
    db().commit()
    return redirect(url_for("assignments", site_id=t["site_id"]))


def task_has_contractor_responsible(task_id):
    """True when a responsible person on the task is the main contractor (site role קבלן)."""
    return db().execute(
        """SELECT 1 FROM task_assignments a JOIN tasks t ON t.id = a.task_id
           JOIN memberships m ON m.contact_id = a.contact_id AND m.site_id = t.site_id
           WHERE a.task_id=? AND a.is_responsible=1 AND m.site_role IN ('קבלן', 'קבלן ראשי')""", (task_id,)).fetchone() is not None


@app.route("/tasks/<int:task_id>/assignees/<int:contact_id>/responsible/add", methods=["POST"])
def add_responsible(task_id, contact_id):
    """With a main contractor responsible, everyone listed as responsible gets the reminder."""
    t = db().execute("SELECT site_id FROM tasks WHERE id=?", (task_id,)).fetchone() or abort(404)
    row = db().execute("SELECT confirmed, is_cc FROM task_assignments WHERE task_id=? AND contact_id=?",
                       (task_id, contact_id)).fetchone() or abort(404)
    if not row["confirmed"] or row["is_cc"] or not task_has_contractor_responsible(task_id):
        abort(400)
    db().execute("UPDATE task_assignments SET is_responsible=1 WHERE task_id=? AND contact_id=?",
                 (task_id, contact_id))
    db().commit()
    return redirect(url_for("assignments", site_id=t["site_id"]))


@app.route("/tasks/<int:task_id>/assignees/<int:contact_id>/cc", methods=["POST"])
def toggle_cc(task_id, contact_id):
    """Mark a confirmed person on a task as copied (מכותב): kept for reports, never sent a reminder."""
    t = db().execute("SELECT site_id FROM tasks WHERE id=?", (task_id,)).fetchone() or abort(404)
    row = db().execute("SELECT is_cc, is_responsible FROM task_assignments WHERE task_id=? AND contact_id=?",
                       (task_id, contact_id)).fetchone() or abort(404)
    if row["is_responsible"]:
        abort(400)  # the responsible party cannot also be a copy
    db().execute("UPDATE task_assignments SET is_cc=? WHERE task_id=? AND contact_id=?",
                 (0 if row["is_cc"] else 1, task_id, contact_id))
    db().commit()
    return redirect(url_for("assignments", site_id=t["site_id"]))


@app.route("/sites/<int:site_id>/cc-report")
def cc_report(site_id):
    """How often each person was copied on tasks, and which tasks."""
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    rows = db().execute(
        """SELECT c.first_name, c.last_name, c.role, c.org, t.public_no, t.topic, t.text
           FROM task_assignments a JOIN tasks t ON t.id = a.task_id
           JOIN contacts c ON c.id = a.contact_id
           WHERE t.site_id=? AND a.is_cc=1 ORDER BY c.first_name, c.last_name, t.public_no""",
        (site_id,)).fetchall()
    people = {}
    for r in rows:
        key = (r["first_name"], r["last_name"] or "")
        people.setdefault(key, {"name": f"{r['first_name']} {r['last_name'] or ''}".strip(),
                                "role": r["role"] or "", "org": r["org"] or "", "tasks": []})
        people[key]["tasks"].append(f"{short_no(r['public_no'])} ({r['topic'] or 'ללא נושא'})")
    return page("דוח מכותבים - " + s["name"], """
<p class="note">מספר המשימות שבהן איש הקשר סומן כמכותב. המכותבים לא מקבלים תזכורות.</p>
<table><tr><th>איש קשר</th><th>תפקיד</th><th>ארגון</th><th>מספר פעמים</th><th>משימות</th></tr>
{% for p in people.values()|sort(attribute='name') %}<tr><td>{{ p['name'] }}</td><td>{{ p['role'] }}</td>
<td>{{ p['org'] }}</td><td>{{ p['tasks']|length }}</td><td>{{ p['tasks']|join(' · ') }}</td></tr>{% else %}
<tr><td colspan="5">אין מכותבים מסומנים עדיין</td></tr>{% endfor %}</table>
<p><a href="{{ url_for('assignments', site_id=s['id']) }}">חזרה לשיוך משימות</a></p>
""", people=people, s=s, url_for=url_for)


@app.route("/sites/<int:site_id>/assignments/rematch", methods=["POST"])
def assignments_rematch(site_id):
    db().execute("SELECT 1 FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    rematch_site(site_id)
    return redirect(url_for("assignments", site_id=site_id))


@app.route("/sites/<int:site_id>/assignments")
def assignments(site_id):
    rows = db().execute(
        """SELECT t.id AS task_id, t.public_no, t.topic, t.text, t.assignee_text, a.contact_id, a.raw_name,
                  a.confirmed, a.is_responsible, a.is_cc, c.first_name, c.last_name, c.role, c.org,
                  EXISTS(SELECT 1 FROM task_assignments x JOIN memberships m
                         ON m.contact_id = x.contact_id AND m.site_id = t.site_id
                         WHERE x.task_id = t.id AND x.is_responsible = 1 AND m.site_role IN ('קבלן', 'קבלן ראשי')) AS contractor_task
           FROM tasks t JOIN task_assignments a ON a.task_id = t.id
           JOIN contacts c ON c.id = a.contact_id
           WHERE t.site_id=? ORDER BY t.public_no, a.is_responsible DESC""", (site_id,)).fetchall()
    return page("שיוך משימות לאנשי קשר", """
<p class="note">הקישור נעשה לפי שם שמופיע בעמודת 'לטיפול'. המפקח מאשר כל שיוך לפני שהוא נכנס לתור ההודעות.
בכל משימה יש אדם אחד אחראי לביצוע, והאחרים מקבלים גישה לצפייה בלבד.</p>
<table><tr><th>משימה</th><th>תיאור המשימה</th><th>טקסט בפרוטוקול</th><th>הציע</th><th>איש קשר</th><th>תפקיד בשיוך</th><th>פעולה</th></tr>
{% for r in rows %}<tr><td>{{ r['public_no']|pubno }}</td><td style="min-width:260px">{{ r['topic'] or '' }}: {{ r['text'] }}</td><td>{{ r['assignee_text'] }}</td><td>{{ r['raw_name'] }}</td>
<td><a href="{{ url_for('contact_card', site_id=site_id, contact_id=r['contact_id']) }}">{{ r['first_name'] }} {{ r['last_name'] or '' }}</a> ({{ r['role'] or '' }}, {{ r['org'] or '' }})</td>
<td>{% if not r['confirmed'] %}ממתין לאישור{% elif r['is_responsible'] %}אחראי לביצוע{% elif r['is_cc'] %}מכותב{% else %}לצפייה{% endif %}</td>
<td>{% if not r['confirmed'] %}
<form method="post" action="{{ url_for('confirm_assignee', task_id=r['task_id'], contact_id=r['contact_id']) }}">
<button>אישור השיוך</button></form>
{% else %}{% if not r['is_responsible'] %}
<form method="post" action="{{ url_for('set_responsible', task_id=r['task_id'], contact_id=r['contact_id']) }}">
<button>הפוך לאחראי לביצוע</button></form>
{% if r['contractor_task'] and not r['is_cc'] and not r['is_responsible'] %}
<form method="post" action="{{ url_for('add_responsible', task_id=r['task_id'], contact_id=r['contact_id']) }}">
<button>הוסף כאחראי (קבלן ראשי)</button></form>{% endif %}
<form method="post" action="{{ url_for('toggle_cc', task_id=r['task_id'], contact_id=r['contact_id']) }}">
<button>{% if r['is_cc'] %}בטל מכותב{% else %}סמן כמכותב{% endif %}</button></form>{% endif %}{% endif %}</td></tr>{% endfor %}</table>
<p><a href="{{ url_for('cc_report', site_id=site_id) }}">דוח מכותבים</a></p>
<form method="post" action="{{ url_for('assignments_rematch', site_id=site_id) }}">
<button>חיפוש שיוך מחדש לפי השמות בכל המשימות הפתוחות</button></form>
""", rows=rows, url_for=url_for, site_id=site_id)


def search_terms(q):
    """Words, or quoted phrases ("ליאור ישראלי"). Every term must appear in the task."""
    terms = [m.group(1) or m.group(2) for m in re.finditer(r'"([^"]+)"|(\S+)', q or "")]
    return [t.strip() for t in terms if t and t.strip()]


def highlight(value, terms):
    text = str(value or "")
    if not terms:
        return escape(text)
    pattern = "|".join(re.escape(t) for t in sorted(terms, key=len, reverse=True))
    out, last = [], 0
    for m in re.finditer(pattern, text, flags=re.IGNORECASE):
        out.append(escape(text[last:m.start()]))
        out.append(Markup("<mark>") + escape(m.group(0)) + Markup("</mark>"))
        last = m.end()
    out.append(escape(text[last:]))
    return Markup("").join(out)


@app.route("/sites/<int:site_id>/search")
def site_search(site_id):
    """Search and filter: tasks or decisions; by words, status, contact, topic; quick views."""
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    q = (request.args.get("q") or "").strip()
    kind = request.args.get("kind") if request.args.get("kind") in ("משימות", "החלטות") else "משימות"
    view = request.args.get("view") if request.args.get("view") in ("open", "done", "info", "repeat3", "awaiting", "aged", "repeat2") else ""
    status = request.args.get("status") or ""
    topic = request.args.get("topic") or ""
    contact_id = request.args.get("contact", type=int)
    terms = search_terms(q)
    active = bool(terms or status or topic or contact_id or view)

    statuses = [r[0] for r in db().execute(
        "SELECT DISTINCT status FROM tasks WHERE site_id=? AND status IS NOT NULL ORDER BY status", (site_id,)).fetchall()]
    topics = sorted({r[0] for r in db().execute(
        "SELECT topic FROM tasks WHERE site_id=? AND topic IS NOT NULL UNION SELECT topic FROM decisions "
        "WHERE site_id=? AND topic IS NOT NULL", (site_id, site_id)).fetchall() if r[0]})
    contacts = _full_names(site_id)

    tasks, decisions = [], []
    if kind == "משימות" and (active or not q):
        linked = set()
        if contact_id:
            linked = {r[0] for r in db().execute(
                "SELECT task_id FROM task_assignments WHERE contact_id=?", (contact_id,)).fetchall()}
        open_names = task_status_names("open")
        for t in db().execute("SELECT * FROM tasks WHERE site_id=? ORDER BY public_no", (site_id,)).fetchall():
            if view == "open" and t["status"] in CLOSED_STATUSES + INFO_STATUSES:
                continue
            if view == "awaiting" and t["status"] != "ממתינה לאישור":
                continue
            if view == "aged" and not (t["status"] in open_names and aged_over_30(t)):
                continue
            if view == "repeat2" and not (t["status"] in open_names and task_repeat_count(t["id"]) >= 2):
                continue
            if view == "done" and t["status"] != "בוצע":
                continue
            if view == "info" and t["status"] != "לידיעה":
                continue
            if view == "repeat3" and task_repeat_count(t["id"]) < 3:
                continue
            if status and t["status"] != status:
                continue
            if topic and (t["topic"] or "") != topic:
                continue
            if contact_id and t["id"] not in linked:
                continue
            haystack = " ".join(str(t[k] or "") for k in ("public_no", "topic", "text", "assignee_text",
                                                          "status", "target_text")).lower()
            if all(term.lower() in haystack for term in terms):
                tasks.append(t)
    if kind == "החלטות":
        for d in db().execute("SELECT * FROM decisions WHERE site_id=? ORDER BY decided_on DESC, id DESC",
                              (site_id,)).fetchall():
            if topic and (d["topic"] or "") != topic:
                continue
            haystack = " ".join(str(d[k] or "") for k in ("topic", "text", "source", "decided_on")).lower()
            if all(term.lower() in haystack for term in terms):
                decisions.append(d)

    sort = request.args.get("sort", "")
    direction = "desc" if request.args.get("dir") == "desc" else "asc"
    task_keys = {
        "number": lambda t: (parse_date((t["public_no"] or "").split("#")[0]) or "", t["public_no"] or ""),
        "topic": lambda t: (t["topic"] or "").lower(),
        "text": lambda t: (t["text"] or "").lower(),
        "assignee": lambda t: (t["assignee_text"] or "").lower(),
        "target": lambda t: (t["target_date"] or "9999-99-99", t["target_text"] or ""),
        "status": lambda t: t["status"] or "",
        "print": lambda t: t["print_flag"] or 0,
    }
    dec_keys = {
        "topic": lambda d: (d["topic"] or "").lower(),
        "text": lambda d: (d["text"] or "").lower(),
        "date": lambda d: parse_date(d["decided_on"]) or (d["decided_on"] or ""),
        "source": lambda d: d["source"] or "",
    }
    if kind == "משימות" and sort in task_keys:
        tasks.sort(key=task_keys[sort], reverse=(direction == "desc"))
    if kind == "החלטות" and sort in dec_keys:
        decisions.sort(key=dec_keys[sort], reverse=(direction == "desc"))

    def heads(columns):
        """Column headers that sort: the same column again reverses the order."""
        args = {k: v for k, v in request.args.items() if k not in ("sort", "dir")}
        out = []
        for label, key in columns:
            nd = "desc" if (sort == key and direction == "asc") else "asc"
            arrow = ("▲" if direction == "asc" else "▼") if sort == key else ""
            out.append({"label": label, "arrow": arrow,
                        "url": url_for("site_search", site_id=s["id"], sort=key, dir=nd, **args)})
        return out
    task_heads = heads([("#", "number"), ("נושא", "topic"), ("משימה", "text"), ("לטיפול (כפי שכתוב)", "assignee"),
                        ("יעד", "target"), ("סטטוס", "status"), ("הדפסה", "print")])
    dec_heads = heads([("נושא", "topic"), ("החלטה", "text"), ("תאריך", "date"), ("מקור", "source")])

    shown = tasks if kind == "משימות" else decisions
    return page("חיפוש וחיתוך - " + s["name"], """
<p>
<a href="{{ url_for('site_search', site_id=s['id'], view='open') }}">משימות פתוחות</a> ·
<a href="{{ url_for('site_search', site_id=s['id'], view='done') }}">משימות שבוצעו</a> ·
<a href="{{ url_for('site_search', site_id=s['id'], view='info') }}">לידיעה</a> ·
<a href="{{ url_for('site_search', site_id=s['id'], kind='החלטות') }}">כל ההחלטות</a> ·
<a href="{{ url_for('site_search', site_id=s['id'], view='repeat3') }}">חזרו 3 פעמים ומעלה</a> ·
<a href="{{ url_for('site_protocol_search', site_id=s['id']) }}">חיפוש בכל הפרוטוקולים</a> ·
<a href="{{ url_for('site_warnings', site_id=s['id']) }}">ללא שיוך</a>
</p>
<form method="get">
<input name="q" value="{{ q }}" size="30" placeholder='מילה, או צמד מילים בגרשיים'>
סוג <select name="kind">
<option {% if kind == 'משימות' %}selected{% endif %}>משימות</option>
<option {% if kind == 'החלטות' %}selected{% endif %}>החלטות</option></select>
{% if kind == 'משימות' %}
סטטוס <select name="status"><option value="">הכול</option>
{% for st in statuses %}<option {% if st == status %}selected{% endif %}>{{ st }}</option>{% endfor %}</select>
איש קשר <select name="contact"><option value="">הכול</option>
{% for c in contacts %}<option value="{{ c['id'] }}" {% if c['id'] == contact_id %}selected{% endif %}>{{ c['first_name'] }} {{ c['last_name'] or '' }}</option>{% endfor %}</select>
{% endif %}
נושא <select name="topic"><option value="">הכול</option>
{% for tp in topics %}<option {% if tp == topic %}selected{% endif %}>{{ tp }}</option>{% endfor %}</select>
<input type="hidden" name="view" value="{{ view }}">
<input type="hidden" name="sort" value="{{ sort }}"><input type="hidden" name="dir" value="{{ direction }}">
<button>חיפוש</button>
</form>
{% if kind == 'החלטות' %}<p class="note">להחלטות אין שיוך לאנשי קשר, ולכן הסינון לפי סטטוס ואיש קשר לא חל עליהן.</p>{% endif %}
{% if shown or active or kind == 'החלטות' %}<p class="note">נמצאו {{ shown|length }} {{ kind }}.</p>{% endif %}
{% if kind == 'משימות' and (active or not q) %}
<table><tr>{% for h in task_heads %}<th><a href="{{ h['url'] }}">{{ h['label'] }} {{ h['arrow'] }}</a></th>{% endfor %}</tr>
{% for t in tasks %}<tr><td>{{ hl(t['public_no']|pubno, terms) }}</td><td>{{ hl(t['topic'], terms) }}</td>
<td>{{ hl(t['text'], terms) }}</td><td>{{ hl(t['assignee_text'], terms) }}</td>
<td>{{ hl(t['target_text'], terms) }}</td><td><span class="pill {{ sclass.get(t['status'], '') }}" role="button" tabindex="0" title="שינוי סטטוס" onclick="this.nextElementSibling.showModal()">{{ hl(t['status'], terms) }}</span><dialog style="text-align:right;min-width:260px">
<form method="post" action="{{ url_for('task_status', task_id=t['id']) }}" data-ajax>
<p><b>{{ t['public_no']|pubno }}</b> {{ t['text'] }}</p>
<p><select name="status" style="width:100%">{% for o in status_options %}<option {% if o == t['status'] %}selected{% endif %}>{{ o }}</option>{% endfor %}</select></p>
<p><input name="by" placeholder="מי טיפל" style="width:100%"></p>
<p><button>עדכון</button> <button type="button" onclick="this.closest('dialog').close()">ביטול</button></p>
</form></dialog></td>
<td><form method="post" action="{{ url_for('task_print_flag', task_id=t['id']) }}">
<input type="hidden" name="next" value="{{ request.full_path }}">
<input type="checkbox" {% if t['print_flag'] %}checked{% endif %} onchange="this.form.submit()" title="סמן להדפסה"></form></td></tr>
{% else %}<tr><td colspan="7">אין תוצאות</td></tr>{% endfor %}</table>
{% endif %}
{% if kind == 'החלטות' %}
<table><tr>{% for h in dec_heads %}<th><a href="{{ h['url'] }}">{{ h['label'] }} {{ h['arrow'] }}</a></th>{% endfor %}</tr>
{% for d in decisions %}<tr><td>{{ hl(d['topic'], terms) }}</td><td>{{ hl(d['text'], terms) }}</td>
<td>{{ d['decided_on'] or '' }}</td><td>{{ d['source'] or '' }}</td></tr>
{% else %}<tr><td colspan="4">אין החלטות תואמות</td></tr>{% endfor %}</table>
{% endif %}
<p><a href="{{ url_for('site_view', site_id=s['id']) }}">חזרה לאתר</a></p>
""", q=q, kind=kind, view=view, status=status, topic=topic, contact_id=contact_id, terms=terms,
        active=active, statuses=statuses, topics=topics, contacts=contacts, tasks=tasks, decisions=decisions,
        shown=shown, hl=highlight, s=s, url_for=url_for, sort=sort, direction=direction,
        task_heads=task_heads, dec_heads=dec_heads, sclass=STATUS_CLASS, status_options=task_status_names())


@app.route("/sites/<int:site_id>/audit")
def site_audit(site_id):
    """Audit of published protocols: the same protocol recorded twice, and continuity per kind.
    A weekly protocol is expected every 7 days; other kinds only show the gap."""
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    meetings = db().execute(
        "SELECT * FROM meetings WHERE site_id=? AND status='published' ORDER BY id", (site_id,)).fetchall()
    items = []
    for m in meetings:
        rows = [r for r in json.loads(m["staged_rows"] or "[]") if not r.get("skip")]
        dates = Counter(parse_date(r.get("meeting_date")) for r in rows if r.get("meeting_date"))
        day = dates.most_common(1)[0][0] if dates else None
        content = sorted([str(r.get("number") or ""), r.get("topic") or "", r.get("task") or "", r.get("target") or ""]
                         for r in rows)
        items.append({"m": m, "day": day, "rows": len(rows), "gap": None, "flags": [],
                      "fp": hashlib.sha1(json.dumps(content, ensure_ascii=False).encode()).hexdigest()})

    # Duplicates: same kind and date, or the same content under another meeting.
    for i, a in enumerate(items):
        if a["day"] is None:
            a["flags"].append("תאריך לא זוהה")
        for b in items[:i]:
            if a["day"] and b["day"] == a["day"] and b["m"]["kind"] == a["m"]["kind"]:
                a["flags"].append(f"כפילות תאריך עם פגישה #{b['m']['id']}")
                break
        for b in items[:i]:
            if a["rows"] and b["fp"] == a["fp"]:
                a["flags"].append(f"תוכן זהה לפגישה #{b['m']['id']}")
                break

    # Continuity per kind, in date order.
    by_kind = {}
    for it in items:
        by_kind.setdefault(it["m"]["kind"], []).append(it)
    for kind, its in by_kind.items():
        its.sort(key=lambda it: (it["day"] or "9999-99-99", it["m"]["id"]))
        prev = None
        for it in its:
            if prev and prev["day"] and it["day"]:
                gap = (dt.date.fromisoformat(it["day"]) - dt.date.fromisoformat(prev["day"])).days
                it["gap"] = gap
                if kind == "שבועית" and gap != 7:
                    it["flags"].append(f"פער של {gap} ימים (צפוי 7)")
            prev = it

    items.sort(key=lambda it: (it["day"] or "9999-99-99", it["m"]["id"]))
    problems = sum(1 for it in items if it["flags"])
    return page("ביקורת פרוטוקולים - " + s["name"], """
<p><a href="{{ url_for('site_view', site_id=s['id']) }}">חזרה לאתר</a></p>
<p class="note">נבדקו {{ items|length }} פרוטוקולים שפורסמו. עם בעיה: {{ problems }}.
המחזור השבועי צפוי כל 7 ימים. בסוגים אחרים מוצג רק הפער בין הפגישות.</p>
<table><tr><th>סוג</th><th>תאריך</th><th>מזהה</th><th>שם קובץ</th><th>שורות</th><th>ימים מהקודם</th><th>בדיקה</th><th>פעולות</th></tr>
{% for it in items %}<tr{% if it['flags'] %} class="overdue"{% endif %}><td>{{ it['m']['kind'] }}</td><td>{{ it['day'] or '' }}</td>
<td>#{{ it['m']['id'] }}</td><td>{{ it['m']['source_name'] or '' }}</td><td>{{ it['rows'] }}</td>
<td>{{ it['gap'] if it['gap'] is not none else '' }}</td><td>{{ it['flags']|join(' · ') if it['flags'] else 'תקין' }}</td>
<td><a href="{{ url_for('meeting_view', meeting_id=it['m']['id']) }}">פתיחה</a>
<form method="post" action="{{ url_for('meeting_update', meeting_id=it['m']['id']) }}" style="display:inline">
<input type="date" name="meeting_date" value="{{ it['m']['meeting_date'] or '' }}"> <select name="kind"><option{% if it['m']['kind']=='ביצוע' %} selected{% endif %}>ביצוע</option><option{% if it['m']['kind']=='שבועית' %} selected{% endif %}>שבועית</option><option{% if it['m']['kind']=='סופרפוזיציה' %} selected{% endif %}>סופרפוזיציה</option></select> <button>שמירה</button></form>
<form method="post" action="{{ url_for('meeting_delete', meeting_id=it['m']['id']) }}" style="display:inline" onsubmit="return confirm('למחוק את הפרוטוקול ואת המשימות שנוצרו רק ממנו? הפעולה אינה הפיכה.')">
<input type="hidden" name="confirm" value="מחק"><button>מחיקה</button></form></td></tr>
{% else %}<tr><td colspan="8">אין פרוטוקולים שפורסמו</td></tr>{% endfor %}</table>
""", s=s, items=items, problems=problems)


@app.route("/meetings/<int:meeting_id>/update", methods=["POST"])
def meeting_update(meeting_id):
    """Correct the protocol date and kind. Tasks keep their own dates; only the protocol record changes."""
    m = db().execute("SELECT * FROM meetings WHERE id=?", (meeting_id,)).fetchone() or abort(404)
    try:
        day = dt.date.fromisoformat(request.form.get("meeting_date", "")).isoformat()
    except ValueError:
        day = m["meeting_date"]
    kind = request.form.get("kind") or m["kind"]
    db().execute("UPDATE meetings SET meeting_date=?, kind=? WHERE id=?", (day, kind, meeting_id))
    db().commit()
    return redirect(url_for("site_audit", site_id=m["site_id"]))


@app.route("/meetings/<int:meeting_id>/delete", methods=["POST"])
def meeting_delete(meeting_id):
    """Delete a protocol and the tasks that came only from it (with what was recorded against them)."""
    m = db().execute("SELECT * FROM meetings WHERE id=?", (meeting_id,)).fetchone() or abort(404)
    if request.form.get("confirm") != "מחק":
        abort(400)
    d = db()
    task_ids = [r[0] for r in d.execute(
        "SELECT DISTINCT task_id FROM task_occurrences WHERE meeting_id=?", (meeting_id,)).fetchall()]
    d.execute("DELETE FROM task_occurrences WHERE meeting_id=?", (meeting_id,))
    d.execute("DELETE FROM task_images WHERE meeting_id=?", (meeting_id,))
    d.execute("DELETE FROM task_targets WHERE meeting_id=?", (meeting_id,))
    d.execute("DELETE FROM task_closure_pending WHERE meeting_id=?", (meeting_id,))
    d.execute("DELETE FROM admin_queue WHERE meeting_id=?", (meeting_id,))
    for tid in task_ids:
        if d.execute("SELECT 1 FROM task_occurrences WHERE task_id=?", (tid,)).fetchone() is None:
            for table in ("task_status_log", "task_comments", "task_docs", "task_links", "task_assignments", "task_images"):
                d.execute(f"DELETE FROM {table} WHERE task_id=?", (tid,))
            d.execute("DELETE FROM tasks WHERE id=?", (tid,))
    d.execute("DELETE FROM meetings WHERE id=?", (meeting_id,))
    d.commit()
    return redirect(url_for("site_audit", site_id=m["site_id"]))


@app.route("/tasks/<int:task_id>/closure/<action>", methods=["POST"])
def task_closure(task_id, action):
    """A task that dropped out of a full protocol waits for the supervisor: close it or keep it open."""
    if action not in ("approve", "keep"):
        abort(400)
    t = db().execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone() or abort(404)
    if db().execute("SELECT 1 FROM task_closure_pending WHERE task_id=?", (task_id,)).fetchone():
        db().execute("DELETE FROM task_closure_pending WHERE task_id=?", (task_id,))
        if action == "approve":
            db().execute("UPDATE tasks SET status='בוצע', closed_at=? WHERE id=?", (dt.date.today().isoformat(), task_id))
            log_task_status(task_id, "בוצע", "נסגרה לאחר היעלמה מפרוטוקול מלא")
        db().commit()
    return _task_reply(task_id, url_for("site_warnings", site_id=t["site_id"]), hide=True,
                       message="הסגירה נשמרה." if action == "approve" else "המשימה נשארה פתוחה.")


@app.route("/tasks/<int:task_id>/comments", methods=["POST"])
def task_comment_add(task_id):
    t = db().execute("SELECT site_id FROM tasks WHERE id=?", (task_id,)).fetchone() or abort(404)
    text = (request.form.get("text") or "").strip()
    if text:
        me = current_contact()
        db().execute("INSERT INTO task_comments(task_id, contact_id, text, created_at) VALUES (?,?,?,?)",
                     (task_id, me["id"] if me else None, text, dt.datetime.now().isoformat(timespec="minutes")))
        db().commit()
    return redirect(url_for("site_view", site_id=t["site_id"]))


@app.route("/tasks/<int:task_id>/docs", methods=["POST"])
def task_doc_add(task_id):
    t = db().execute("SELECT site_id FROM tasks WHERE id=?", (task_id,)).fetchone() or abort(404)
    name = (request.form.get("name") or "").strip()
    if name:
        db().execute("INSERT INTO task_docs(task_id, name, created_at) VALUES (?,?,?)",
                     (task_id, name, dt.datetime.now().isoformat(timespec="minutes")))
        db().commit()
    return redirect(url_for("site_view", site_id=t["site_id"]))


@app.route("/u/<token>", methods=["GET", "POST"])
def task_link_view(token):
    """Status update for one task, reached from a link. No login: the link itself is the access."""
    row = db().execute("SELECT * FROM task_links WHERE token_hash=?",
                       (hashlib.sha256(token.encode()).hexdigest(),)).fetchone()
    if row is None or row["expires_at"] < dt.datetime.now().isoformat(timespec="minutes"):
        abort(404)
    t = db().execute("SELECT * FROM tasks WHERE id=?", (row["task_id"],)).fetchone() or abort(404)
    c = db().execute("SELECT * FROM contacts WHERE id=?", (row["contact_id"],)).fetchone() or abort(404)
    name = f"{c['first_name'] or ''} {c['last_name'] or ''}".strip() or c["email"]
    allowed = [s for s in task_status_names() if s not in LINK_STATUSES_EXCLUDED]
    message = ""
    if request.method == "POST":
        new_status = request.form.get("status", "")
        comment = (request.form.get("comment") or "").strip()
        if new_status == "בוצע":
            db().execute("INSERT OR IGNORE INTO task_closure_pending(task_id, meeting_id, created_at) VALUES (?,?,?)",
                         (t["id"], t["last_meeting_id"], dt.datetime.now().isoformat(timespec="minutes")))
            message = "בקשת הסגירה נשלחה למפקח לאישור."
        elif new_status in allowed:
            db().execute("UPDATE tasks SET status=? WHERE id=?", (new_status, t["id"]))
            db().execute(
                "INSERT INTO task_status_log(task_id, status, handled_by, changed_by, changed_at) VALUES (?,?,?,?,?)",
                (t["id"], new_status, name, name, dt.datetime.now().isoformat(timespec="minutes")))
            message = "הסטטוס עודכן."
            t = db().execute("SELECT * FROM tasks WHERE id=?", (t["id"],)).fetchone()
            notify_supervisors_task_handled(t, name, new_status)
        elif new_status != "בוצע":
            abort(400)
        doc = (request.form.get("doc") or "").strip()
        if doc:
            db().execute("INSERT INTO task_docs(task_id, name, created_at) VALUES (?,?,?)",
                         (t["id"], doc, dt.datetime.now().isoformat(timespec="minutes")))
        if comment:
            db().execute("INSERT INTO task_comments(task_id, contact_id, text, created_at) VALUES (?,?,?,?)",
                         (t["id"], c["id"], comment, dt.datetime.now().isoformat(timespec="minutes")))
        db().commit()
        t = db().execute("SELECT * FROM tasks WHERE id=?", (t["id"],)).fetchone()
        if _wants_json():
            return jsonify(ok=True, status=t["status"], css=STATUS_CLASS.get(t["status"], ""), message=message or "העדכון נשמר.")
    return page("עדכון משימה", """
<h2>עדכון משימה {{ t['public_no']|pubno }}</h2>
<p><b>{{ t['topic'] or '' }}</b></p><p>{{ t['text'] }}</p>
<p class="note">סטטוס נוכחי: <span id="status-now">{{ t['status'] }}</span></p>
{% if message %}<p class="flash">{{ message }}</p>{% endif %}
<form method="post" data-ajax>
<label>סטטוס חדש <select name="status">{% for o in allowed %}<option>{{ o }}</option>{% endfor %}<option>בוצע</option></select></label>
<p class="note">בחירה ב"בוצע" שולחת בקשת סגירה למפקח; המשימה תיסגר רק לאחר אישורו.</p>
<label>אסמכתא (לא חובה: שם מסמך או קישור)<br><input name="doc" style="width:100%"></label>
<label>הערה (לא חובה)<br><textarea name="comment" rows="3" style="width:100%"></textarea></label>
<p><button>שמירה</button></p>
</form>
""", t=t, allowed=allowed, message=message)


@app.route("/sites/<int:site_id>/send-links", methods=["GET", "POST"])
def site_send_links(site_id):
    """Send each responsible person a personal link per task, by the channel chosen in the table."""
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    me = current_contact()
    if not (me["is_admin"] or "מפקח" in roles_on_site(me, site_id)):
        abort(403)
    rows = db().execute(
        """SELECT a.contact_id, c.first_name, c.last_name, c.email, t.id AS task_id, t.public_no, t.topic, t.text
           FROM task_assignments a JOIN contacts c ON c.id = a.contact_id JOIN tasks t ON t.id = a.task_id
           WHERE t.site_id=? AND a.is_responsible=1 AND a.confirmed=1
             AND t.status NOT IN ('בוצע', 'לידיעה', 'בוטלה')
           ORDER BY c.first_name, c.last_name, t.public_no""", (site_id,)).fetchall()
    sent = None
    if request.method == "POST":
        groups = {}
        for r in rows:
            ch = request.form.get(f"ch_{r['task_id']}_{r['contact_id']}", "email")
            if ch == "none":
                continue
            groups.setdefault((r["contact_id"], ch), []).append(r)
        base = request.host_url.rstrip("/")
        messages = links = 0
        for (contact_id, ch), items in groups.items():
            c = db().execute("SELECT * FROM contacts WHERE id=?", (contact_id,)).fetchone()
            lines = [f"שלום {c['first_name'] or ''}, נא לעדכן סטטוס למשימות הבאות:"]
            for it in items:
                token = make_task_link(it["task_id"], contact_id)
                links += 1
                lines.append(f"{short_no(it['public_no'])} – {it['topic'] or ''}: {(it['text'] or '')[:120]}")
                lines.append(f"לעדכון: {base}/u/{token}")
            enqueue(site_id, c, "\n".join(lines), "מיידי", channel=ch)
            messages += 1
        db().commit()
        sent = (messages, links)
    return page("שליחת קישורי עדכון - " + s["name"], """
<h2>שליחת קישורי עדכון לאחראים</h2>
<p class="note">כל אחראי מקבל קישור אישי לכל משימה. הקישור פתוח לשבועיים ומאפשר לעדכן רק את המשימה שלו.</p>
{% if sent %}<p class="flash">נוצרו {{ sent[0] }} הודעות ו-{{ sent[1] }} קישורים.</p>{% endif %}
<form method="post">
<table><tr><th>אחראי</th><th>משימה</th><th>נושא</th><th>שליחה</th></tr>
{% for r in rows %}<tr><td>{{ r['first_name'] }} {{ r['last_name'] or '' }}</td><td>{{ r['public_no']|pubno }}</td><td>{{ r['topic'] or '' }}: {{ r['text'][:80] }}</td>
<td><select name="ch_{{ r['task_id'] }}_{{ r['contact_id'] }}"><option value="email">דוא"ל</option><option value="mobile">נייד</option><option value="none">לא לשלוח</option></select></td></tr>
{% else %}<tr><td colspan="4">אין משימות פתוחות עם אחראי מאושר</td></tr>{% endfor %}</table>
<p><button>שליחה</button></p>
</form>
<p class="note">במצב פיתוח כל ההודעות נשלחות לבעלים בלבד, עם שם הנמען המקורי בגוף ההודעה.</p>
""", rows=rows, sent=sent, s=s)


@app.route("/sites/<int:site_id>/protocol-search")
def site_protocol_search(site_id):
    """Search every row of every protocol, including rows that are not tasks."""
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    q = request.args.get("q", "")
    terms = [x.lower() for x in search_terms(q)]
    hits = []
    if terms:
        for m in db().execute("SELECT * FROM meetings WHERE site_id=? ORDER BY id", (site_id,)).fetchall():
            for r in json.loads(m["staged_rows"] or "[]"):
                text = " ".join(str(r.get(k) or "") for k in ("number", "topic", "task", "target", "meeting_date", "subject"))
                if all(term in text.lower() for term in terms):
                    hits.append({"kind": m["kind"], "meeting": m["id"], "row": r, "skip": bool(r.get("skip"))})
    return page("חיפוש בפרוטוקולים - " + s["name"], """
<p><a href="{{ url_for('site_search', site_id=s['id']) }}">חזרה לחיפוש</a></p>
<form method="get"><input name="q" value="{{ q }}" placeholder="מילה או ביטוי"> <button>חיפוש</button></form>
{% if terms %}<p class="note">{{ hits|length }} תוצאות (עד 300 מוצגות)</p>
<table><tr><th>סוג</th><th>ישיבה</th><th>תאריך</th><th>סעיף</th><th>נושא</th><th>טקסט</th><th>יעד</th></tr>
{% for h in hits[:300] %}<tr><td>{{ h['kind'] }}</td><td>#{{ h['meeting'] }}</td><td>{{ h['row'].get('meeting_date') or '' }}</td>
<td>{{ h['row'].get('number') or '' }}</td><td>{{ h['row'].get('topic') or '' }}</td><td>{{ h['row'].get('task') or '' }}</td><td>{{ h['row'].get('target') or '' }}</td></tr>{% endfor %}</table>
{% endif %}
""", s=s, q=q, terms=terms, hits=hits)


@app.route("/sites/<int:site_id>/warnings")
def site_warnings(site_id):
    """Things to fix: the same person recorded more than once, tasks with no one assigned, and
    names that still wait for the supervisor. The name check is run again first."""
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    rematch_site(site_id)
    people = db().execute(
        """SELECT c.id, c.first_name, c.last_name, c.email, c.org, m.site_role
           FROM memberships m JOIN contacts c ON c.id = m.contact_id WHERE m.site_id=?
           ORDER BY c.first_name, c.last_name, c.id""", (site_id,)).fetchall()
    groups = {}
    for p in people:
        key = ((p["first_name"] or "").strip(), (p["last_name"] or "").strip())
        if key[0] and key[1]:
            groups.setdefault(key, {})
            groups[key].setdefault(p["id"], {"email": p["email"], "org": p["org"] or "", "roles": []})
            groups[key][p["id"]]["roles"].append(p["site_role"])
    duplicates = [{"name": f"{k[0]} {k[1]}", "records": list(v.values())}
                  for k, v in groups.items() if len(v) > 1]
    unassigned = db().execute(
        """SELECT t.id, t.public_no, t.topic, t.text, t.assignee_text FROM tasks t
           WHERE t.site_id=? AND t.status NOT IN ('בוצע', 'לידיעה', 'בוטלה')
             AND NOT EXISTS (SELECT 1 FROM task_assignments a WHERE a.task_id=t.id)
           ORDER BY t.public_no""", (site_id,)).fetchall()
    pending = db().execute(
        """SELECT t.id AS task_id, t.public_no, t.topic, t.text, a.raw_name, a.contact_id, c.first_name, c.last_name, c.org
           FROM task_assignments a JOIN tasks t ON t.id = a.task_id JOIN contacts c ON c.id = a.contact_id
           WHERE t.site_id=? AND t.status NOT IN ('בוצע', 'בוטלה') AND a.confirmed=0 ORDER BY t.public_no""", (site_id,)).fetchall()
    choices = db().execute(
        """SELECT DISTINCT c.id, c.first_name, c.last_name FROM memberships m JOIN contacts c ON c.id = m.contact_id
           WHERE m.site_id=? ORDER BY c.first_name, c.last_name""", (site_id,)).fetchall()
    closures = db().execute(
        """SELECT t.id, t.public_no, t.topic, t.text FROM task_closure_pending p
           JOIN tasks t ON t.id = p.task_id WHERE t.site_id=? ORDER BY t.public_no""", (site_id,)).fetchall()
    return page("התרעות - " + s["name"], """
<h3>אנשי קשר כפולים (אותו שם, יותר מרשומה אחת) - מומלץ למחוק את הכפילות</h3>
<table><tr><th>שם</th><th>רשומות</th></tr>
{% for d in duplicates %}<tr><td>{{ d['name'] }}</td><td>{% for r in d['records'] %}{{ r['email'] if show_email else '' }} ({{ r['org'] }}: {{ r['roles']|join(', ') }})<br>{% endfor %}</td></tr>
{% else %}<tr><td colspan="2">אין כפילויות</td></tr>{% endfor %}</table>
<h3>לבדיקה: לא הופיעה בפרוטוקול המלא האחרון (לא נסגרה אוטומטית)</h3>
<table><tr><th>משימה</th><th>תיאור</th><th>פעולה</th></tr>
{% for c in closures %}<tr><td>{{ c['public_no']|pubno }}</td><td>{{ c['topic'] or '' }}: {{ c['text'] }}</td>
<td><form method="post" data-ajax action="{{ url_for('task_closure', task_id=c['id'], action='approve') }}" style="display:inline"><button>סגור כבוצע</button></form>
<form method="post" data-ajax action="{{ url_for('task_closure', task_id=c['id'], action='keep') }}" style="display:inline"><button>השאר פתוחה</button></form></td></tr>
{% else %}<tr><td colspan="3">אין סגירות ממתינות</td></tr>{% endfor %}</table>
<h3>שיוכים הממתינים לאישור המפקח (שם חלקי או שם משותף לכמה אנשים)</h3>
<table><tr><th>משימה</th><th>תיאור המשימה</th><th>הטקסט בפרוטוקול</th><th>איש קשר מוצע</th><th>ארגון</th></tr>
{% for p in pending %}<tr><td>{{ p['public_no']|pubno }}</td><td style="min-width:260px">{{ p['topic'] or '' }}: {{ p['text'] }}</td>
<td>{{ p['raw_name'] }}</td>
<td><form method="post" action="{{ url_for('assign_task', task_id=p['task_id']) }}">
<select name="contact_id">{% for c in choices %}<option value="{{ c['id'] }}" {% if c['id'] == p['contact_id'] %}selected{% endif %}>{{ c['first_name'] }} {{ c['last_name'] or '' }}</option>{% endfor %}</select>
<button>שייך</button></form></td><td>{{ p['org'] or '' }}</td></tr>
{% else %}<tr><td colspan="5">אין שיוכים ממתינים. לאישור השיוכים: <a href="{{ url_for('assignments', site_id=s['id']) }}">שיוך משימות</a></td></tr>{% endfor %}</table>
<h3>משימות פתוחות ללא שיוך לאיש קשר</h3>
<table><tr><th>#</th><th>נושא</th><th>משימה</th><th>לטיפול (כפי שכתוב)</th></tr>
{% for t in unassigned %}<tr><td>{{ t['public_no']|pubno }}</td><td>{{ t['topic'] or '' }}</td><td>{{ t['text'] }}</td>
<td>{{ t['assignee_text'] or '' }}
<form method="post" data-ajax action="{{ url_for('assign_task', task_id=t['id']) }}" style="margin-top:4px">
<select name="contact_id">{% for c in choices %}<option value="{{ c['id'] }}">{{ c['first_name'] }} {{ c['last_name'] or '' }}</option>{% endfor %}</select>
<button>שייך</button></form>
<form method="post" data-ajax action="{{ url_for('task_triage', task_id=t['id']) }}" style="margin-top:4px">
<button name="action" value="info">1 - לידיעה</button>
<button name="action" value="delete" onclick="return confirm('למחוק את המשימה? הפעולה אינה הפיכה.')">2 - מחק</button></form></td></tr>
{% else %}<tr><td colspan="4">כל המשימות הפתוחות משויכות</td></tr>{% endfor %}</table>
<p><a href="{{ url_for('site_view', site_id=s['id']) }}">חזרה לאתר</a></p>
""", duplicates=duplicates, unassigned=unassigned, pending=pending, choices=choices, closures=closures,
        show_email=can_see_emails(site_id), s=s, url_for=url_for)


@app.route("/sites/<int:site_id>/contacts/<int:contact_id>")
def contact_card(site_id, contact_id):
    """Personal details card of a contact on this site. Read-only."""
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    c = db().execute("SELECT id, first_name, last_name, org, role, phone, mobile, email, address FROM contacts WHERE id=?",
                     (contact_id,)).fetchone() or abort(404)
    roles = [r[0] for r in db().execute(
        "SELECT site_role FROM memberships WHERE site_id=? AND contact_id=? ORDER BY site_role",
        (site_id, contact_id)).fetchall()]
    tasks = db().execute(
        """SELECT t.public_no, t.topic, t.status, a.confirmed, a.is_responsible, a.is_cc FROM task_assignments a
           JOIN tasks t ON t.id = a.task_id WHERE t.site_id=? AND a.contact_id=? ORDER BY t.public_no""",
        (site_id, contact_id)).fetchall()
    return page("כרטיס איש קשר", """
<h3>{{ c['first_name'] }} {{ c['last_name'] or '' }}</h3>
<table>
<tr><th>ארגון</th><td>{{ c['org'] or '' }}</td></tr>
<tr><th>תפקיד</th><td>{{ c['role'] or '' }}</td></tr>
<tr><th>תפקיד באתר</th><td>{{ roles|join(', ') }}</td></tr>
<tr><th>טלפון</th><td>{{ c['phone'] or '' }}</td></tr>
<tr><th>נייד</th><td>{{ c['mobile'] or '' }}</td></tr>
{% if show_email %}<tr><th>דוא"ל</th><td dir="ltr">{{ c['email'] or '' }}</td></tr>{% endif %}
<tr><th>כתובת</th><td>{{ c['address'] or '' }}</td></tr>
</table>
<h3>משימות</h3>
<table><tr><th>משימה</th><th>סטטוס</th><th>תפקיד בשיוך</th></tr>
{% for t in tasks %}<tr><td>{{ t['public_no']|pubno }} {{ t['topic'] or '' }}</td><td>{{ t['status'] }}</td>
<td>{% if not t['confirmed'] %}ממתין לאישור{% elif t['is_responsible'] %}אחראי לביצוע{% elif t['is_cc'] %}מכותב{% else %}לצפייה{% endif %}</td></tr>
{% else %}<tr><td colspan="3">אין משימות משויכות</td></tr>{% endfor %}</table>
<p><a href="{{ url_for('site_contacts', site_id=s['id']) }}">חזרה לרשימת אנשי הקשר</a></p>
""", c=c, roles=roles, tasks=tasks, show_email=can_see_emails(site_id), s=s, url_for=url_for)


@app.route("/sites/<int:site_id>/contacts")
def site_contacts(site_id):
    """Read-only list of the project's contacts, with the site roles each one holds."""
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    rows = db().execute(
        """SELECT c.id, c.first_name, c.last_name, c.org, c.phone, c.mobile, c.email,
                  GROUP_CONCAT(m.site_role, ', ') AS site_roles
           FROM contacts c JOIN memberships m ON m.contact_id = c.id
           WHERE m.site_id=? GROUP BY c.id ORDER BY c.first_name, c.last_name""", (site_id,)).fetchall()
    return page("אנשי קשר - " + s["name"], """
{% if can_edit %}<p class="note">מנהל המערכת והמפקח יכולים לערוך פרטים. מחיקת שדות - למנהל המערכת בלבד.</p>
<p><a href="{{ url_for('contacts_import') }}">ייבוא אנשי קשר מקובץ</a></p>
{% else %}<p class="note">רשימה לצפייה בלבד.</p>{% endif %}
{% if is_admin %}<form method="post" action="{{ url_for('contacts_bulk_delete') }}" onsubmit="return confirm('למחוק את האנשים שסומנו? הפעולה אינה הפיכה.')">{% endif %}
<table><tr>{% if is_admin %}<th><input type="checkbox" onclick="document.querySelectorAll('.pick').forEach(function(c){c.checked=this.checked}, this)" title="בחירת כל הרשומות"></th>{% endif %}
<th>שם</th><th>ארגון</th><th>תפקיד באתר</th><th>טלפון</th><th>נייד</th>{% if show_email %}<th>דוא"ל</th>{% endif %}{% if can_edit %}<th></th>{% endif %}</tr>
{% for r in rows %}<tr>{% if is_admin %}<td><input type="checkbox" class="pick" name="ids" value="{{ r['id'] }}"></td>{% endif %}
<td>{{ r['first_name'] }} {{ r['last_name'] or '' }}</td><td>{{ r['org'] or '' }}</td>
<td>{{ r['site_roles'] }}</td><td>{{ r['phone'] or '' }}</td><td>{{ r['mobile'] or '' }}</td>{% if show_email %}<td>{{ r['email'] or '' }}</td>{% endif %}
{% if can_edit %}<td><a href="{{ url_for('contact_edit', contact_id=r['id']) }}">עריכה</a></td>{% endif %}</tr>{% endfor %}</table>
{% if is_admin %}<input type="hidden" name="site_id" value="{{ s['id'] }}"><p><button>מחיקת הנבחרים</button> <span class="note">המחיקה מוחקת את האדם מכל האתרים. מנהל המערכת לא נמחק.</span></p>{% endif %}
{% if is_admin %}</form>{% endif %}
<p><a href="{{ url_for('site_view', site_id=s['id']) }}">חזרה לאתר</a></p>
""", rows=rows, s=s, url_for=url_for, can_edit=is_manager(current_contact()), show_email=can_see_emails(site_id),
        is_admin=bool(current_contact()["is_admin"]))


@app.route("/contacts/bulk-delete", methods=["POST"])
def contacts_bulk_delete():
    """Administrator only (not listed in SECTION_OF, so require_access blocks everyone else)."""
    me = current_contact()
    site_id = request.form.get("site_id", type=int)
    ids = [int(x) for x in request.form.getlist("ids") if x.isdigit()]
    for cid in ids:
        if cid == me["id"]:
            continue
        row = db().execute("SELECT is_admin FROM contacts WHERE id=?", (cid,)).fetchone()
        if row is None or row["is_admin"]:
            continue
        db().execute("DELETE FROM task_assignments WHERE contact_id=?", (cid,))
        db().execute("DELETE FROM memberships WHERE contact_id=?", (cid,))
        db().execute("DELETE FROM contacts WHERE id=?", (cid,))
    db().commit()
    if site_id:
        return redirect(url_for("site_contacts", site_id=site_id))
    return redirect(url_for("sites"))

# ---------- routes ----------
PAGE = """<!doctype html><html lang="he" dir="rtl"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{ title }}</title>
<style>
:root{--red:#C62828;--red-light:#FBE9E9;--white:#FFFFFF;--gray:#6B6B6B;--gray-light:#F2F2F2;--gray-line:#D0D0D0}
body{font-family:Arial,sans-serif;margin:0;background:var(--white);color:#222;line-height:1.5}
header{background:var(--red);color:var(--white);padding:12px 16px}
header a{color:var(--white);margin-left:16px;text-decoration:none;font-weight:bold}
main{padding:16px;max-width:1100px;margin:0 auto}
h2{color:var(--red);border-bottom:2px solid var(--red);padding-bottom:4px}
h3{color:var(--gray);margin-top:24px}
table{border-collapse:collapse;width:100%;margin-top:8px}
td,th{border:1px solid var(--gray-line);padding:6px;vertical-align:top;font-size:14px}
th{background:var(--gray-light);color:var(--gray);text-align:right}
tr.overdue{background:var(--red-light);color:var(--red);font-weight:bold}
tr.urgent{border-right:4px solid var(--red)}
tr.done{background:var(--gray-light);color:var(--gray)}
button,input[type=submit]{background:var(--red);color:var(--white);border:0;padding:6px 14px;border-radius:3px;cursor:pointer}
button:hover{background:#A31F1F}
input,select,textarea{border:1px solid var(--gray-line);padding:4px;border-radius:3px}
.note{color:var(--gray);font-size:13px}
.flash{background:var(--red-light);color:var(--red);padding:8px;border-radius:3px}
@media (max-width:640px){table.cards,table.cards tbody,table.cards tr,table.cards td{display:block;width:100%}table.cards tr:first-child{display:none}table.cards tr{border:1px solid #ddd;border-radius:6px;margin:0 0 10px;padding:6px;background:#fff}table.cards td{border:0;padding:3px 0;min-width:0!important;white-space:normal!important}table.cards td:empty{display:none}table.cards td::before{content:attr(data-label);display:block;color:#6B6B6B;font-size:12px}}
.pill{display:inline-block;padding:1px 8px;border-radius:10px;font-size:12px;background:#E9ECEF;color:#333;white-space:nowrap}
.s-open{background:#E3F0FF;color:#1B4F8C}.s-prog{background:#E0F5F1;color:#0E6B5B}.s-appr{background:#FFF4D6;color:#8A5A00}
.s-spec{background:#EFE7FB;color:#5B2C91}.s-recheck{background:#FDE7D2;color:#A04A00}.s-review{background:#DDF3F7;color:#11657A}
.s-hold{background:#E4E4E4;color:#444}.s-info{background:#F4F4F4;color:#666}.s-done{background:#DFF5E3;color:#1D6B33}
.s-cancel{background:#F0F0F0;color:#888;text-decoration:line-through}
.hero{margin:0 0 12px;border-radius:6px;overflow:hidden}.hero img{width:100%;height:260px;object-fit:cover;display:block}
.kpis{display:flex;flex-wrap:wrap;gap:10px;margin:10px 0}.kpi{flex:1 1 150px;border:1px solid var(--gray-line);border-radius:6px;padding:10px;background:#fff;display:block;color:inherit;text-decoration:none}a.kpi:hover{border-color:#999}
.kpi b{display:block;font-size:26px;color:#222}.kpi.warn b{color:var(--red)}.kpi span{color:var(--gray);font-size:13px}
.brow{display:flex;align-items:center;gap:8px;margin:4px 0;font-size:14px}.brow .lbl{width:160px}
.brow .track{flex:1;background:var(--gray-light);height:12px;border-radius:6px;overflow:hidden}
.brow .fill{display:block;height:100%;background:#7A8699}.brow .n{width:40px;text-align:left}
.stepper{display:flex;flex-wrap:wrap;gap:6px;margin:8px 0 14px}.step{border:1px solid var(--gray-line);border-radius:14px;padding:3px 10px;font-size:13px;color:var(--gray);background:#fff}
.step.done{background:var(--gray-light);color:#444}.step.active{background:var(--red);border-color:var(--red);color:#fff;font-weight:bold}
.pill[role=button]{cursor:pointer;text-decoration:underline dotted}
dialog{border:1px solid var(--gray-line);border-radius:8px;padding:16px;box-shadow:0 4px 16px rgba(0,0,0,.2)}
.shell{display:flex;flex-direction:column;align-items:stretch}
.side{flex:none;background:#F4F4F4;border-bottom:1px solid #E0E0E0;padding:6px 10px;display:flex;flex-wrap:wrap;align-items:center;gap:2px}
.side a{display:block;padding:7px 12px;color:#333;text-decoration:none;font-size:14px;border-radius:4px}
.side a:hover{background:#E9E9E9}
.side a.on{background:var(--red);color:#fff;font-weight:bold}
.side-title{padding:4px 12px;color:var(--red);font-weight:bold;font-size:14px;margin-left:8px}
.side-sep{border-top:1px solid #DDD;margin:8px 0}
main{flex:1;min-width:0}
@media (max-width:640px){header{padding:8px 10px;display:flex;align-items:center;flex-wrap:nowrap;gap:10px}header b{font-size:15px;white-space:nowrap}header a{margin-left:0;font-size:13px;white-space:nowrap}header{overflow-x:auto}html,body{max-width:100vw;overflow-x:hidden}.shell{flex-direction:column;width:100%;min-width:0}.shell>main{min-width:0;max-width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch}.side{flex:none;min-height:0;display:flex;flex-wrap:nowrap;overflow-x:auto;white-space:nowrap;padding:6px 8px;gap:4px;-webkit-overflow-scrolling:touch}.side a{padding:7px 12px;font-size:14px;flex:0 0 auto;border-radius:16px;background:#fff;border:1px solid #E0E0E0}.side a.on,.side a[aria-current]{background:var(--red);color:#fff;border-color:var(--red)}.side-title,.side-sep{display:none}}
</style></head><body>
<header><b>הלל הזקן 14</b>
{% if home_url %}<a href="{{ home_url }}">לוח בקרה</a>{% endif %}<a href="{{ url_for('sites') }}">אתרים</a><a href="{{ url_for('contacts_import') }}">ייבוא אנשי קשר</a></header>
<div class="shell">{% if side or side_admin %}<aside class="side">{% if side_title %}<div class="side-title">{{ side_title }}</div>{% endif %}
{% for it in side %}<a href="{{ it['url'] }}" class="{{ 'on' if it['on'] else '' }}">{{ it['label'] }}</a>{% endfor %}
{% if side_admin %}<div class="side-sep"></div>{% for it in side_admin %}<a href="{{ it['url'] }}" class="{{ 'on' if it['on'] else '' }}">{{ it['label'] }}</a>{% endfor %}{% endif %}
</aside>{% endif %}
<main><h2>{{ title }}</h2>
{% for m in get_flashed_messages() %}<p class="flash">{{ m }}</p>{% endfor %}
{{ body|safe }}</main></div>
<script>
(function(){
  function toast(text, bad){
    var t=document.getElementById('ajax-toast');
    if(!t){ t=document.createElement('div'); t.id='ajax-toast';
      t.style.cssText='position:fixed;bottom:16px;left:16px;right:16px;max-width:360px;margin:auto;padding:10px 14px;border-radius:6px;color:#fff;z-index:9999;font-size:14px;box-shadow:0 2px 8px rgba(0,0,0,.2)';
      document.body.appendChild(t); }
    t.style.background = bad ? '#C62828' : '#2E7D32';
    t.textContent = text; t.style.display = 'block';
    clearTimeout(t._h); t._h = setTimeout(function(){ t.style.display = 'none'; }, 3500);
  }
  function applyResult(form, d){
    var dlg = form.closest('dialog');
    if (dlg) { dlg.close(); }
    var row = form.closest('tr');
    if (d.hide && row) {
      row.style.transition = 'opacity .3s'; row.style.opacity = '0';
      setTimeout(function(){ row.remove(); }, 300);
    } else if (row && d.status) {
      var pill = row.querySelector('.pill');
      if (pill) { pill.textContent = d.status; pill.className = 'pill ' + (d.css || ''); }
    }
    var now = document.getElementById('status-now');
    if (now && d.status) { now.textContent = d.status; }
  }
  document.addEventListener('submit', function(e){
    var f = e.target;
    if (!f.matches || !f.matches('form[data-ajax]')) return;
    e.preventDefault();
    var fd = new FormData(f);
    if (e.submitter && e.submitter.name) { fd.append(e.submitter.name, e.submitter.value); }
    var buttons = f.querySelectorAll('button');
    buttons.forEach(function(b){ b.disabled = true; });
    fetch(f.action || window.location.href, {method: 'POST', body: fd, credentials: 'same-origin',
        headers: {'X-Requested-With': 'fetch', 'Accept': 'application/json'}})
      .then(function(r){ if (!r.ok) throw new Error('העדכון לא נשמר, נסו שוב'); return r.json(); })
      .then(function(d){ applyResult(f, d); toast(d.message || 'העדכון נשמר', false); })
      .catch(function(err){ toast(err.message || 'העדכון לא נשמר', true); })
      .finally(function(){ buttons.forEach(function(b){ b.disabled = false; }); });
  });
})();
</script>
</body></html>"""


SIDE_SITE_ITEMS = [
    ("לוח בקרה", "site_view", "tasks"),
    ("משימות וחיפוש", "site_search", "tasks"),
    ("ללא שיוך", "site_warnings", "tasks"),
    ("שיוך משימות", "assignments", "tasks"),
    ("אנשי קשר של הפרויקט", "site_contacts", "tasks"),
    ("שליחת קישורי עדכון", "site_send_links", "tasks"),
    ("להדפסה", "site_print", "tasks"),
    ("מאגר פרוטוקולים", "site_protocols", "protocols"),
    ("העלאת פרוטוקול", "meeting_upload", "protocols"),
    ("ביקורת פרוטוקולים", "site_audit", "protocols"),
]


def _side_menu():
    """Side menu for signed-in people: the site's screens they may open, then general administration."""
    if not has_request_context():
        return {}
    me = current_contact()
    if me is None or request.endpoint is None or request.endpoint in PUBLIC_ENDPOINTS:
        return {}
    site_id = (request.view_args or {}).get("site_id")
    side, title = [], ""
    if site_id is not None:
        row = db().execute("SELECT name FROM sites WHERE id=?", (site_id,)).fetchone()
        if row is not None:
            title = row["name"]
            sections = set()
            if not me["is_admin"]:
                for role in roles_on_site(me, site_id):
                    sections |= role_sections(role)
            for label, endpoint, section in SIDE_SITE_ITEMS:
                if me["is_admin"] or section in sections:
                    side.append({"label": label, "url": url_for(endpoint, site_id=site_id),
                                 "on": request.endpoint == endpoint})
    admin = []
    if me["is_admin"]:
        admin = [{"label": "ניהול כללי", "url": url_for("admin_overview"), "on": request.endpoint == "admin_overview"},
                 {"label": "אתר חדש", "url": url_for("site_new"), "on": request.endpoint == "site_new"},
                 {"label": "כל האתרים", "url": url_for("sites"), "on": request.endpoint == "sites"}]
    home_url = ""
    if side:
        home_url = url_for("site_view", site_id=site_id)
    else:
        first = db().execute("SELECT id FROM sites ORDER BY id LIMIT 1").fetchone()
        if first is not None:
            home_url = url_for("site_view", site_id=first["id"])
    return {"side": side, "side_admin": admin, "side_title": title, "home_url": home_url}


def page(title, body, **ctx):
    return render_template_string(PAGE, title=title, body=render_template_string(body, **ctx),
                                  url_for=url_for, get_flashed_messages=lambda: [], **_side_menu())


@app.route("/")
def sites():
    rows = []
    for r in db().execute("SELECT * FROM sites ORDER BY id").fetchall():
        d = dict(r)
        end_txt = construction_schedule(d["land_date"], d["months_total"])[0]
        d["land_shown"] = _dmy(d["land_date"])
        d["delivery_shown"] = _dmy(d["delivery_date"]) if d["delivery_date"] else (
            f"{end_txt} (מחושב)" if end_txt else "")
        rows.append(d)
    return page("אתרים", """
<p><a href="{{ url_for('site_new') }}">אתר חדש</a> · <a href="{{ url_for('admin_overview') }}">ניהול כללי</a></p>
<table><tr><th>אתר</th><th>עלייה לקרקע</th><th>מסירה</th><th>חודשי ביצוע (חוזה)</th><th></th></tr>
{% for s in rows %}<tr><td>{{ s['name'] }}</td><td>{{ s['land_shown'] }}</td><td>{{ s['delivery_shown'] }}</td>
<td>{{ s['months_total'] }}</td><td><a href="{{ url_for('site_view', site_id=s['id']) }}">פתיחה</a></td></tr>{% endfor %}</table>
""", rows=rows, url_for=url_for)


@app.route("/sites/new", methods=["GET", "POST"])
def site_new():
    if request.method == "POST":
        f = request.form
        db().execute(
            "INSERT INTO sites(name, address, land_date, delivery_date, months_total, contract_ref) VALUES (?,?,?,?,?,?)",
            (f["name"], f.get("address", ""), parse_date(f.get("land_date", "")),
             parse_date(f.get("delivery_date", "")), int(f.get("months_total") or 0), f.get("contract_ref", "")))
        db().commit()
        return redirect(url_for("sites"))
    return page("אתר חדש", """
<form method="post">
<p>שם פרויקט <input name="name" required></p><p>כתובת <input name="address"></p>
<p>תאריך עלייה לקרקע (לפי החוזה, dd.mm.yy) <input name="land_date"></p>
<p>תאריך מסירה (לפי החוזה) <input name="delivery_date"></p>
<p>סך חודשי הביצוע (לפי החוזה) <input name="months_total" type="number"></p>
<p>מזהה חוזה <input name="contract_ref"></p>
<button>שמירה</button></form>""", url_for=url_for)


def _import_dir():
    return os.path.join(UPLOAD_DIR, "imports")


def _valid_token(token):
    return len(token) == 32 and all(c in "0123456789abcdef" for c in token)


def _desired_roles(role_text):
    parts = [_site_role_of(p) for p in (role_text or "").split(",")]
    return [r for r in dict.fromkeys(p for p in parts if p)]


def _role_plan(rows, site_id):
    """Compare the file with the site's roles, person by person. People not in the file keep
    their roles; a row with no role leaves that person's roles alone."""
    desired_by_contact, plan = {}, {}
    for r in rows:
        email = (r.get("email") or "").strip().lower()
        if not email or not re.fullmatch(EMAIL_PATTERN, email):
            continue
        c = db().execute("SELECT id, first_name, last_name FROM contacts WHERE email=?", (email,)).fetchone()
        if c is None:
            plan.setdefault(email, {"email": email, "name": "", "contact_id": None,
                                    "status": "לא קיים במערכת - ידולג"})
            continue
        desired = desired_by_contact.setdefault(c["id"], set())
        desired.update(_desired_roles(r.get("role")))
        plan[email] = {"email": email, "name": f"{c['first_name'] or ''} {c['last_name'] or ''}".strip(),
                       "contact_id": c["id"], "status": ""}
    out = []
    for item in plan.values():
        cid = item["contact_id"]
        if cid is None:
            out.append(item)
            continue
        current = {x[0] for x in db().execute(
            "SELECT site_role FROM memberships WHERE site_id=? AND contact_id=?", (site_id, cid)).fetchall()}
        desired = desired_by_contact.get(cid, set())
        if not desired:
            item["status"] = "אין תפקיד בקובץ - לא משתנה"
            item.update(current=sorted(current), add=[], remove=[])
        else:
            item.update(current=sorted(current), add=sorted(desired - current), remove=sorted(current - desired),
                        desired=sorted(desired))
            item["status"] = "ללא שינוי" if not (desired ^ current) else "ייעודכן"
        out.append(item)
    return out


@app.route("/roles/sync", methods=["GET"])
def roles_sync():
    all_sites = db().execute("SELECT id, name FROM sites ORDER BY id").fetchall()
    return page("סנכרון תפקידים לפי קובץ (למנהל המערכת)", """
<p class="note">לכל איש קשר שמופיע בקובץ, התפקידים באתר יהיו בדיוק אלה שבקובץ. אנשים שלא בקובץ לא ישתנו. לפני האישור תראה מה יתווסף ומה יוסר.</p>
<form method="post" action="{{ url_for('roles_sync_preview') }}" enctype="multipart/form-data">
<p>אתר <select name="site_id">{% for s in all_sites %}<option value="{{ s['id'] }}">{{ s['name'] }}</option>{% endfor %}</select></p>
<p>קובץ xlsx <input type="file" name="file" accept=".xlsx" required></p>
<button>בדיקה</button></form>
""", all_sites=all_sites, url_for=url_for)


@app.route("/roles/sync/preview", methods=["POST"])
def roles_sync_preview():
    f = request.files.get("file")
    site_id = request.form.get("site_id", type=int)
    if not f or not site_id:
        flash("יש לבחור אתר וקובץ")
        return redirect(url_for("roles_sync"))
    os.makedirs(_import_dir(), exist_ok=True)
    token = uuid.uuid4().hex
    tmp_xlsx = os.path.join(_import_dir(), token + ".xlsx")
    f.save(tmp_xlsx)
    try:
        rows = _contact_rows(tmp_xlsx)
    except ValueError as exc:
        os.remove(tmp_xlsx)
        flash(str(exc))
        return redirect(url_for("roles_sync"))
    os.remove(tmp_xlsx)
    with open(_import_rows_file(token), "w", encoding="utf-8") as fh:
        json.dump(rows, fh, ensure_ascii=False)
    plan = _role_plan(rows, site_id)
    site = db().execute("SELECT name FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    return page("בדיקת סנכרון תפקידים", """
<p>אתר: <b>{{ site_name }}</b>. זו תצוגה מקדימה בלבד: עדיין לא נשמר דבר.</p>
<table><tr><th>דוא"ל</th><th>שם</th><th>תפקידים היום</th><th>בקובץ</th><th>יתווסף</th><th>יוסר</th><th>סטטוס</th></tr>
{% for p in plan %}<tr><td dir="ltr">{{ p['email'] }}</td><td>{{ p['name'] }}</td>
<td>{{ (p['current'] or [])|join(', ') }}</td><td>{{ (p['desired'] or [])|join(', ') }}</td>
<td>{{ (p['add'] or [])|join(', ') }}</td><td>{{ (p['remove'] or [])|join(', ') }}</td><td>{{ p['status'] }}</td></tr>{% endfor %}</table>
<form method="post" action="{{ url_for('roles_sync_confirm') }}">
<input type="hidden" name="token" value="{{ token }}"><input type="hidden" name="site_id" value="{{ site_id }}">
<button>אישור הסנכרון</button></form>
""", plan=plan, token=token, site_id=site_id, site_name=site["name"], url_for=url_for)


@app.route("/roles/sync/confirm", methods=["POST"])
def roles_sync_confirm():
    token = request.form.get("token", "")
    site_id = request.form.get("site_id", type=int)
    path = _import_rows_file(token) if _valid_token(token) else None
    if not site_id or not path or not os.path.exists(path):
        flash("הבדיקה פגה, העלה את הקובץ שוב")
        return redirect(url_for("roles_sync"))
    with open(path, encoding="utf-8") as fh:
        rows = json.load(fh)
    os.remove(path)
    for item in _role_plan(rows, site_id):
        if item.get("contact_id") is None or item["status"] != "ייעודכן":
            continue
        for role in item["remove"]:
            db().execute("DELETE FROM memberships WHERE site_id=? AND contact_id=? AND site_role=?",
                         (site_id, item["contact_id"], role))
        for role in item["add"]:
            db().execute("INSERT OR IGNORE INTO memberships(site_id, contact_id, site_role) VALUES (?,?,?)",
                         (site_id, item["contact_id"], role))
    db().commit()
    rematch_site(site_id)
    return redirect(url_for("site_view", site_id=site_id))


@app.route("/contacts/<int:contact_id>/edit", methods=["GET", "POST"])
def contact_edit(contact_id):
    c = db().execute("SELECT * FROM contacts WHERE id=?", (contact_id,)).fetchone() or abort(404)
    me = current_contact()
    values = {key: (c[key] or "") for key, _ in CONTACT_EDIT_FIELDS}
    error = None
    if request.method == "POST":
        for key, _ in CONTACT_EDIT_FIELDS:
            new = (request.form.get(key) or "").strip()
            if not new and not me["is_admin"]:
                new = values[key]  # supervisors cannot blank a field; only the administrator can
            values[key] = new
        email = values["email"]
        if email and not re.fullmatch(EMAIL_PATTERN, email):
            error = 'כתובת הדוא"ל לא תקינה.'
        elif email and db().execute("SELECT 1 FROM contacts WHERE email=? AND id<>?",
                                    (email, contact_id)).fetchone():
            error = 'כתובת הדוא"ל כבר שייכת לאיש קשר אחר.'
        if error is None:
            db().execute(
                "UPDATE contacts SET first_name=?, last_name=?, org=?, role=?, phone=?, mobile=?, email=?, address=? "
                "WHERE id=?",
                (values["first_name"] or None, values["last_name"] or None, values["org"] or None,
                 values["role"] or None, values["phone"] or None, values["mobile"] or None,
                 values["email"] or None, values["address"] or None, contact_id))
            db().commit()
            return redirect(url_for("sites"))
    memberships = db().execute(
        """SELECT m.site_id, m.site_role, s.name FROM memberships m JOIN sites s ON s.id = m.site_id
           WHERE m.contact_id=? ORDER BY s.id, m.site_role""", (contact_id,)).fetchall()
    return page("עריכת איש קשר", """
<form method="post">
{% for key, label in fields %}<p>{{ label }} <input name="{{ key }}" value="{{ values[key] }}" size="30"></p>{% endfor %}
<button>שמירה</button></form>
{% if error %}<p class="note">{{ error }}</p>{% endif %}
{% if not is_admin %}<p class="note">שדה ריק משאיר את הערך הקיים. מחיקת שדות מותרת למנהל המערכת בלבד.</p>{% endif %}
<h3>תפקידים באתרים</h3>
<table><tr><th>אתר</th><th>תפקיד</th>{% if is_admin %}<th></th>{% endif %}</tr>
{% for m in memberships %}<tr><td>{{ m['name'] }}</td><td>{{ m['site_role'] }}</td>
{% if is_admin %}<td><form method="post" action="{{ url_for('contact_role_remove', contact_id=cid) }}">
<input type="hidden" name="site_id" value="{{ m['site_id'] }}"><input type="hidden" name="site_role" value="{{ m['site_role'] }}">
<button>הסרה</button></form></td>{% endif %}</tr>{% else %}<tr><td colspan="3">אין תפקידים</td></tr>{% endfor %}</table>
<form method="post" action="{{ url_for('contact_role_add', contact_id=cid) }}">
אתר <select name="site_id">{% for s in sites %}<option value="{{ s['id'] }}">{{ s['name'] }}</option>{% endfor %}</select>
תפקיד <input name="site_role" size="16"> <button>הוספה</button></form>
{% if is_admin and not c['is_admin'] %}
<h3>מחיקת איש קשר</h3>
<form method="post" action="{{ url_for('contact_delete', contact_id=cid) }}">
<button>מחיקת איש הקשר (מנהל המערכת בלבד)</button></form>
{% endif %}
""", fields=CONTACT_EDIT_FIELDS, values=values, error=error, is_admin=bool(me["is_admin"]),
        memberships=memberships, sites=db().execute("SELECT id, name FROM sites ORDER BY id").fetchall(),
        cid=contact_id, c=c)


@app.route("/contacts/<int:contact_id>/roles/add", methods=["POST"])
def contact_role_add(contact_id):
    db().execute("SELECT 1 FROM contacts WHERE id=?", (contact_id,)).fetchone() or abort(404)
    site_id = request.form.get("site_id", type=int)
    role = _site_role_of(request.form.get("site_role", ""))
    if not site_id or not role or not db().execute("SELECT 1 FROM sites WHERE id=?", (site_id,)).fetchone():
        abort(400)
    db().execute("INSERT OR IGNORE INTO memberships(site_id, contact_id, site_role) VALUES (?,?,?)",
                 (site_id, contact_id, role))
    db().commit()
    return redirect(url_for("contact_edit", contact_id=contact_id))


@app.route("/contacts/<int:contact_id>/roles/remove", methods=["POST"])
def contact_role_remove(contact_id):
    if not current_contact()["is_admin"]:
        abort(403)  # removing a role is administrator-only, like clearing fields
    db().execute("DELETE FROM memberships WHERE contact_id=? AND site_id=? AND site_role=?",
                 (contact_id, request.form.get("site_id", type=int), request.form.get("site_role", "")))
    db().commit()
    return redirect(url_for("contact_edit", contact_id=contact_id))


@app.route("/contacts/<int:contact_id>/delete", methods=["POST"])
def contact_delete(contact_id):
    me = current_contact()
    if not me["is_admin"]:
        abort(403)  # deletion is administrator-only
    if contact_id == me["id"]:
        abort(400)  # the administrator cannot delete their own account
    c = db().execute("SELECT is_admin FROM contacts WHERE id=?", (contact_id,)).fetchone() or abort(404)
    if c["is_admin"]:
        abort(400)
    db().execute("DELETE FROM task_assignments WHERE contact_id=?", (contact_id,))
    db().execute("DELETE FROM memberships WHERE contact_id=?", (contact_id,))
    db().execute("DELETE FROM contacts WHERE id=?", (contact_id,))
    db().commit()
    return redirect(url_for("sites"))


@app.route("/contacts/import", methods=["GET"])
def contacts_import():
    all_sites = db().execute("SELECT id, name FROM sites ORDER BY id").fetchall()
    return page("ייבוא אנשי קשר", """
<form method="post" action="{{ url_for('contacts_import_preview') }}" enctype="multipart/form-data">
<p><label>אתר: <select name="site_id" required>
{% for s in sites %}<option value="{{ s['id'] }}">{{ s['name'] }}</option>{% endfor %}</select></label></p>
<p><input type="file" name="file" accept=".xlsx" required></p>
<button>בדיקה לפני ייבוא</button></form>
<p class="note">התפקיד במערכת נלקח מעמודת "תפקיד" (או "סוג") בכל שורה: יזם, נציגות, בעלי דירות, יועצים, מפקח, קבלן. בשלב הבדיקה תראה מה ייובא ומה דורש תיקון, ורק אחרי אישור הייבוא יתבצע.</p>""",
                url_for=url_for, sites=all_sites)


def _import_rows_file(token):
    return os.path.join(_import_dir(), token + ".json")


@app.route("/contacts/import/preview", methods=["POST"])
def contacts_import_preview():
    f = request.files.get("file")
    site_id = request.form.get("site_id", type=int)
    if not f or not site_id:
        flash("יש לבחור אתר וקובץ")
        return redirect(url_for("contacts_import"))
    os.makedirs(_import_dir(), exist_ok=True)
    token = uuid.uuid4().hex
    tmp_xlsx = os.path.join(_import_dir(), token + ".xlsx")
    f.save(tmp_xlsx)
    try:
        rows = _contact_rows(tmp_xlsx)
    except ValueError as exc:
        os.remove(tmp_xlsx)
        flash(str(exc))
        return redirect(url_for("contacts_import"))
    os.remove(tmp_xlsx)
    with open(_import_rows_file(token), "w", encoding="utf-8") as fh:
        json.dump(rows, fh, ensure_ascii=False)
    site = db().execute("SELECT name FROM sites WHERE id=?", (site_id,)).fetchone()
    return page("בדיקה לפני ייבוא", """
<p>אתר: <b>{{ site_name }}</b> | שורות בקובץ: {{ rows|length }}. סמן או בטל את הסימון "ייובא" לכל שורה, ערוך כל שדה לפי הצורך, ואז לחץ על אישור.</p>
<form method="post" action="{{ url_for('contacts_import_confirm') }}">
<input type="hidden" name="token" value="{{ token }}"><input type="hidden" name="site_id" value="{{ site_id }}">
<table><tr><th>ייובא</th><th>שורה</th><th>דוא"ל</th><th>שם פרטי</th><th>שם משפחה</th><th>תפקיד</th>
<th>ארגון</th><th>טלפון</th><th>נייד</th><th>כתובת</th><th>הערות</th></tr>
{% for r in rows %}{% set i = loop.index0 %}
<tr style="{% if r['notes'] %}background:#fff4d6{% endif %}">
<td><input type="checkbox" name="include_{{ i }}" {% if r['email'] and not r['notes'] or (r['email'] and r['notes'] and 'מופיע גם' in r['notes'][0]) %}checked{% endif %}></td>
<td>{{ r['line'] }}</td>
<td><input name="email_{{ i }}" value="{{ r['email'] }}" dir="ltr"></td>
<td><input name="first_{{ i }}" value="{{ r['first'] }}"></td>
<td><input name="last_{{ i }}" value="{{ r['last'] }}"></td>
<td><input name="role_{{ i }}" value="{{ r['role'] }}"></td>
<td><input name="org_{{ i }}" value="{{ r['org'] }}"></td>
<td><input name="phone_{{ i }}" value="{{ r['phone'] }}" dir="ltr"></td>
<td><input name="mobile_{{ i }}" value="{{ r['mobile'] }}" dir="ltr"></td>
<td><input name="address_{{ i }}" value="{{ r['address'] }}"></td>
<td>{{ r['notes']|join(' | ') }}</td></tr>{% endfor %}
</table>
<p><button>אישור וייבוא</button> <a href="{{ url_for('contacts_import') }}">ביטול</a></p></form>""",
                url_for=url_for, rows=rows, token=token, site_id=site_id,
                site_name=site["name"] if site else "")


@app.route("/contacts/import/confirm", methods=["POST"])
def contacts_import_confirm():
    token = request.form.get("token", "")
    site_id = request.form.get("site_id", type=int)
    if not _valid_token(token) or not site_id:
        flash("הבדיקה פגה, העלה את הקובץ שוב")
        return redirect(url_for("contacts_import"))
    path = _import_rows_file(token)
    if not os.path.exists(path):
        flash("הקובץ לא נמצא, העלה אותו שוב")
        return redirect(url_for("contacts_import"))
    with open(path, encoding="utf-8") as fh:
        rows = json.load(fh)
    os.remove(path)
    chosen = []
    for i, row in enumerate(rows):
        if request.form.get(f"include_{i}") != "on":
            continue
        edited = dict(row)
        for field in CONTACT_FIELDS:
            value = request.form.get(f"{field}_{i}")
            if value is not None:
                edited[field] = value
        chosen.append(edited)
    added, skipped = import_rows(chosen, site_id=site_id)
    rematch_site(site_id)
    flash(f"יובאו {added} שורות (איחוד לפי דוא\"ל). שורות שדולגו: {skipped}")
    return redirect(url_for("site_view", site_id=site_id))


def _linked_people(site_id):
    """task_id -> [(contact_id, name)] for the site, for linking names to their cards."""
    out = {}
    for r in db().execute(
            """SELECT a.task_id, c.id, c.first_name, c.last_name FROM task_assignments a
               JOIN tasks t ON t.id = a.task_id JOIN contacts c ON c.id = a.contact_id
               WHERE t.site_id=? ORDER BY a.is_responsible DESC, c.first_name""", (site_id,)).fetchall():
        out.setdefault(r[0], []).append({"id": r[1], "name": f"{r[2] or ''} {r[3] or ''}".strip()})
    return out


STATUS_CLASS = {
    "פתוחה": "s-open", "בטיפול": "s-prog", "ממתינה לאישור": "s-appr", "ממתין למפרט": "s-spec",
    "לבדיקה חוזרת": "s-recheck", "ממתינה לבדיקה": "s-review", "מעוכבת": "s-hold",
    "לידיעה": "s-info", "בוצע": "s-done", "בוטלה": "s-cancel",
}
SITE_IMAGES = {"הלל הזקן 14": "site-hilel14.jpg"}  # rendering shown in the site header


def site_live_data(site_id):
    """Figures shown on the site home page; also polled by the page so it stays current."""
    kp = dict(site_kpis(site_id))
    today = dt.date.today().isoformat()
    open_names = task_status_names("open")
    marks = ",".join("?" * len(open_names))
    kp["late"] = db().execute(
        f"SELECT COUNT(*) FROM tasks WHERE site_id=? AND target_date IS NOT NULL AND target_date < ? AND status IN ({marks})",
        (site_id, today, *open_names)).fetchone()[0]
    latest = db().execute(
        "SELECT * FROM meetings WHERE site_id=? AND status='published' ORDER BY meeting_date DESC, published_at DESC, id DESC LIMIT 1",
        (site_id,)).fetchone()
    return kp, latest


@app.route("/sites/<int:site_id>/live")
def site_live(site_id):
    db().execute("SELECT 1 FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    kp, latest = site_live_data(site_id)
    return jsonify({
        "kpis": {k: kp[k] for k in ("open", "awaiting_approval", "aged_over_30", "closed", "late")},
        "latest_meeting_id": latest["id"] if latest else 0,
        "time": dt.datetime.now().strftime("%H:%M:%S"),
    })


def _dmy(iso):
    try:
        return dt.date.fromisoformat(iso).strftime("%d.%m.%Y")
    except (TypeError, ValueError):
        return iso or ""


def construction_schedule(land_iso, total):
    """Last construction month (from the land date and the contract's month count) and the
    construction month we are in today, counting the land month as month 1."""
    try:
        land = dt.date.fromisoformat(land_iso)
        total = int(total or 0)
    except (TypeError, ValueError):
        return None, "לא נקבע"
    if total <= 0:
        return None, "לא נקבע"
    start = land.year * 12 + (land.month - 1)
    end_idx = start + total - 1
    end_txt = f"{end_idx % 12 + 1:02d}.{end_idx // 12}"
    today = dt.date.today()
    idx = today.year * 12 + (today.month - 1) - start
    if idx < 0:
        current = f"טרם החל (עלייה לקרקע ב-{_dmy(land_iso)})"
    elif idx >= total:
        current = f"הסתיימו {total} חודשי הביצוע לפי החוזה"
    else:
        current = f"{idx + 1} מתוך {total}"
    return end_txt, current


@app.route("/sites/<int:site_id>")
def site_view(site_id):
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    tasks = db().execute("SELECT * FROM tasks WHERE site_id=?", (site_id,)).fetchall()
    today = dt.date.today().isoformat()
    sort = request.args.get("sort", "")
    sort_keys = {
        "number": lambda t: t["public_no"] or "",
        "topic": lambda t: (t["topic"] or "", t["public_no"] or ""),
        "target": lambda t: (t["target_date"] or "9999-99-99", t["public_no"] or ""),
        "status": lambda t: (t["status"] or "", t["public_no"] or ""),
    }
    open_t = [t for t in tasks if t["status"] not in CLOSED_STATUSES + INFO_STATUSES]
    if sort == "added":  # newest first; tasks without a recorded date go last
        open_t.sort(key=lambda t: t["public_no"] or "")
        open_t.sort(key=lambda t: t["published_at"] or t["created_at"] or "", reverse=True)
    elif sort in sort_keys:
        open_t.sort(key=sort_keys[sort])
    else:
        open_t.sort(key=lambda t: (priority_of(t, today), t["public_no"]))
    info_t = sorted([t for t in tasks if t["status"] == "לידיעה"], key=lambda t: t["public_no"])
    done_t = [t for t in tasks if t["status"] == "בוצע"]
    month = db().execute("SELECT month_current, month_total, meeting_date FROM meetings WHERE site_id=? AND status='published' AND month_current IS NOT NULL ORDER BY id DESC LIMIT 1", (site_id,)).fetchone()
    closure_ids = {r[0] for r in db().execute(
        "SELECT p.task_id FROM task_closure_pending p JOIN tasks t ON t.id = p.task_id WHERE t.site_id=?",
        (site_id,)).fetchall()}
    last = {}
    for r in db().execute(
            """SELECT l.task_id, l.changed_by, l.changed_at FROM task_status_log l
               JOIN tasks t ON t.id = l.task_id WHERE t.site_id=? ORDER BY l.id""", (site_id,)).fetchall():
        last[r[0]] = {"changed_by": r[1], "changed_at": r[2]}
    kp, latest = site_live_data(site_id)
    total = kp["open"] + kp["closed"]
    pct = round(100 * kp["closed"] / total) if total else 0
    latest_tasks = db().execute(
        "SELECT DISTINCT t.* FROM task_occurrences o JOIN tasks t ON t.id = o.task_id WHERE o.meeting_id=? ORDER BY t.public_no",
        (latest["id"],)).fetchall() if latest else []
    latest_date = ""
    if latest:
        try:
            latest_date = dt.date.fromisoformat(latest["meeting_date"]).strftime("%d.%m.%Y")
        except (TypeError, ValueError):
            latest_date = latest["meeting_date"] or ""
    counts = Counter(t["status"] for t in tasks)
    top = max(counts.values(), default=1)
    status_rows = [{"name": n, "n": c, "w": round(100 * c / top)} for n, c in counts.most_common()]
    site_image = next((url_for("static", filename=f) for k, f in SITE_IMAGES.items() if k in s["name"]), None)
    open_names = task_status_names("open")
    open_now = [t for t in tasks if t["status"] in open_names]

    def age_days(t):
        try:
            return (dt.date.today() - dt.date.fromisoformat((t["created_at"] or "")[:10])).days
        except ValueError:
            return 0
    aged_list = sorted([t for t in open_now if age_days(t) > 90], key=lambda t: -age_days(t))[:4]
    aged_rows = [{"t": t, "days": age_days(t)} for t in aged_list]
    topic_counts = Counter((t["topic"] or "ללא נושא") for t in open_now).most_common(5)
    top_topics = [name for name, _ in topic_counts[:3]]
    topic_top = max((c for _, c in topic_counts), default=1)
    topic_rows = [{"name": n, "n": c, "w": round(100 * c / topic_top)} for n, c in topic_counts]
    return page(s["name"], """
{% if is_admin %}<form method="post" action="{{ url_for('site_guest_link', site_id=s['id']) }}" style="margin:6px 0"><button>קישור צפייה ללא סיסמה לשעה</button></form>{% endif %}
{% if site_image %}<div class="hero"><img src="{{ site_image }}" alt="הדמיית הפרויקט"></div>{% endif %}
<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:8px 0">
<a class="kpi warn" href="{{ url_for('site_view', site_id=s['id'], sort='target') }}" style="border:1px solid #c62828"><b data-live="late" style="color:#c62828">{{ kp.late }}</b><span>משימות באיחור</span></a>
<a class="kpi warn" href="{{ url_for('site_search', site_id=s['id'], view='aged') }}"><b data-live="aged_over_30">{{ kp.aged_over_30 }}</b><span>פתוחות מעל 30 יום</span></a>
<a class="kpi" href="{{ url_for('site_search', site_id=s['id'], view='open') }}"><b data-live="open">{{ kp.open }}</b><span>סה"כ משימות פתוחות</span></a>
<a class="kpi" href="{{ url_for('site_search', site_id=s['id'], view='done') }}"><b data-live="closed">{{ kp.closed }}</b><span>נסגרו</span></a>
</div>
<p class="note">ממתינות לאישור: <a href="{{ url_for('site_search', site_id=s['id'], view='awaiting') }}">{{ kp.awaiting_approval }}</a> · חזרו בפרוטוקולים: <a href="{{ url_for('site_search', site_id=s['id'], view='repeat2') }}">{{ kp.repeats_open }}</a> · אחוז סגירה: {{ pct }}%</p>
<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px;margin:8px 0">
<div class="kpi" style="text-align:right"><b style="font-size:16px">משימות ישנות ופתוחות (מעל 90 יום)</b>
{% for a in aged_rows %}<div style="padding:6px 0;border-top:1px solid #eee"><span style="color:#c62828;font-weight:700">{{ a.days }} יום</span> · {{ a.t['text'][:90] }}</div>{% else %}<div>אין משימות ישנות.</div>{% endfor %}
<p class="note"><a href="{{ url_for('site_search', site_id=s['id'], view='aged') }}">לרשימה המלאה</a></p></div>
<div class="kpi" style="text-align:right"><b style="font-size:16px">נושאים מעכבים (הכי הרבה פתוחים)</b>
{% for t in topic_rows %}<div style="display:flex;gap:8px;align-items:center;padding:4px 0"><b style="width:36px;color:#c62828">{{ t.n }}</b><span style="flex:1;background:#eee;height:10px;border-radius:5px"><span style="display:block;height:10px;border-radius:5px;background:#e57373;width:{{ t.w }}%"></span></span><span style="min-width:90px">{{ t.name }}</span></div>{% else %}<div>אין משימות פתוחות.</div>{% endfor %}</div>
<div class="kpi" style="text-align:right"><b style="font-size:16px">סיכום והמלצות לפעולה</b>
<div style="padding:6px 0"><span style="color:#c62828;font-weight:700">{{ kp.late }}</span> משימות באיחור – נדרש טיפול מיידי מול האחראים</div>
<div style="padding:6px 0"><span style="color:#c62828;font-weight:700">{{ kp.aged_over_30 }}</span> משימות פתוחות מעל 30 יום – מומלץ ישיבת סטטוס ממוקדת</div>
<div style="padding:6px 0">נושאים מרכזיים מעכבים: {{ top_topics|join(', ') or 'אין' }}</div></div>
</div>
<div class="kpis" style="display:none">
<a class="kpi" href="{{ url_for('site_search', site_id=s['id'], view='open') }}"><b data-live="open">{{ kp.open }}</b><span>משימות פתוחות</span></a>
<a class="kpi" href="{{ url_for('site_search', site_id=s['id'], view='awaiting') }}"><b data-live="awaiting_approval">{{ kp.awaiting_approval }}</b><span>ממתינות לאישור</span></a>
<a class="kpi warn" href="{{ url_for('site_search', site_id=s['id'], view='aged') }}"><b data-live="aged_over_30">{{ kp.aged_over_30 }}</b><span>פתוחות מעל 30 יום</span></a>
<a class="kpi warn" href="{{ url_for('site_view', site_id=s['id'], sort='target') }}"><b data-live="late">{{ kp.late }}</b><span>באיחור (מועד היעד עבר)</span></a>
<a class="kpi warn" href="{{ url_for('site_search', site_id=s['id'], view='repeat2') }}"><b>{{ kp.repeats_open }}</b><span>פתוחות שחזרו בפרוטוקולים</span></a>
<a class="kpi" href="{{ url_for('site_search', site_id=s['id'], view='done') }}"><b data-live="closed">{{ kp.closed }}</b><span>נסגרו</span></a>
<div class="kpi"><b>{{ pct }}%</b><span>אחוז סגירה</span></div>
</div>
<p class="note live-strip"><span class="live-dot"></span> חי · עודכן <span id="live-time">{{ now_txt }}</span> · מתרענן כל דקה</p>
<h3>משימות מהפרוטוקול האחרון{% if latest_date %} · {{ latest_date }}{% endif %}</h3>
{% if latest_tasks %}<table class="cards"><tr><th>#</th><th>נושא</th><th>משימה</th><th>יעד</th><th>סטטוס</th></tr>
{% for t in latest_tasks %}<tr><td data-label="#">{{ t['public_no']|pubno }}</td><td data-label="נושא">{{ t['topic'] }}</td><td data-label="משימה" style="width:100%">{{ t['text'] }}</td><td data-label="יעד">{{ t['target_text'] }}</td><td data-label="סטטוס" style="min-width:180px"><span class="pill {{ sclass.get(t['status'], '') }}">{{ t['status'] }}</span>
<form method="post" action="{{ url_for('task_status', task_id=t['id']) }}" data-ajax>
<select name="status" style="width:100%;min-width:150px">{% for o in status_options %}<option {% if o == t['status'] %}selected{% endif %}>{{ o }}</option>{% endfor %}</select>
<input name="by" placeholder="מי טיפל" style="width:100%;min-width:150px;margin-top:4px"><button style="margin-top:4px">עדכון</button></form></td></tr>{% endfor %}</table>
{% elif latest %}<p class="note">בפרוטוקול זה אין משימות מקושרות.</p>{% else %}<p class="note">אין פרוטוקול מפורסם לאתר.</p>{% endif %}
<script>
(function () {
  var url = {{ url_for('site_live', site_id=s['id'])|tojson }};
  var known = {{ latest['id'] if latest else 0 }};
  function tick() {
    fetch(url, {credentials: 'same-origin', headers: {'X-Requested-With': 'fetch'}})
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (d) {
        if (!d) return;
        if (d.latest_meeting_id !== known) { location.reload(); return; }
        document.querySelectorAll('[data-live]').forEach(function (el) {
          var k = el.getAttribute('data-live');
          if (k in d.kpis) el.textContent = d.kpis[k];
        });
        document.getElementById('live-time').textContent = d.time;
      })
      .catch(function () {});
  }
  setInterval(tick, 60000);
})();
</script>
<style>.live-strip{display:flex;align-items:center;gap:6px}.live-dot{width:8px;height:8px;border-radius:50%;background:#2a9d4b;display:inline-block;animation:livepulse 1.6s infinite}@keyframes livepulse{0%{opacity:1}50%{opacity:.25}100%{opacity:1}}</style>
<h3>התפלגות סטטוסים</h3>
<div>{% for r in status_rows %}<div class="brow"><span class="lbl"><span class="pill {{ sclass.get(r.name, '') }}">{{ r.name }}</span></span><span class="track"><span class="fill" style="width:{{ r.w }}%"></span></span><span class="n">{{ r.n }}</span></div>{% endfor %}</div>
<form method="post" action="{{ url_for('site_schedule', site_id=s['id']) }}">
שעת שליחה לסיכומים (ברירת מחדל לכל האתר) <input name="digest_time" value="{{ s['digest_time'] }}" size="5">
יום לסיכום שבועי <select name="weekly_day">{% for d in weekdays %}<option {% if d == s['weekly_day'] %}selected{% endif %}>{{ d }}</option>{% endfor %}</select>
<button>שמירה</button></form>
<p>עלייה לקרקע {{ land_txt }} · מסירה לפי חוזה {{ delivery_txt }} · חודש מסירה מחושב {{ end_txt }} · חודש ביצוע נוכחי {{ month_txt }}</p>
<p><a href="{{ url_for('meeting_upload', site_id=s['id']) }}">העלאת פרוטוקול</a> · <a href="{{ url_for('site_contacts', site_id=s['id']) }}">אנשי קשר של הפרויקט</a> · <a href="{{ url_for('site_warnings', site_id=s['id']) }}">התרעות לתיקון</a> · <a href="{{ url_for('site_search', site_id=s['id']) }}">חיפוש משימות</a> · <a href="{{ url_for('site_print', site_id=s['id']) }}">להדפסה</a> · <a href="{{ url_for('site_audit', site_id=s['id']) }}">ביקורת פרוטוקולים</a> · <a href="{{ url_for('site_send_links', site_id=s['id']) }}">שליחת קישורי עדכון</a> · <a href="{{ url_for('site_protocols', site_id=s['id']) }}">מאגר פרוטוקולים</a></p>
<h3>משימות פתוחות</h3>
<form method="get" style="margin-bottom:8px">מיון לפי <select name="sort">
<option value="" {% if sort == '' %}selected{% endif %}>דחיפות (ברירת מחדל)</option>
<option value="added" {% if sort == 'added' %}selected{% endif %}>זמן פרסום (החדשים קודם)</option>
<option value="number" {% if sort == 'number' %}selected{% endif %}>מספר</option>
<option value="topic" {% if sort == 'topic' %}selected{% endif %}>נושא</option>
<option value="target" {% if sort == 'target' %}selected{% endif %}>יעד</option>
<option value="status" {% if sort == 'status' %}selected{% endif %}>סטטוס</option>
</select> <button>מיין</button></form>
<table class="cards"><tr><th>#</th><th>נושא</th><th>משימה</th><th style="min-width:220px">לטיפול</th><th>יעד</th><th>סטטוס</th><th>הדפסה</th></tr>
{% for t in open_t %}<tr class="{{ rowclass(t) }}"><td data-label="#">{{ t['public_no']|pubno }}</td><td data-label="נושא">{{ t['topic'] }}</td>
<td data-label="משימה" style="width:100%">{{ t['text'] }}</td><td data-label="לטיפול" style="min-width:220px;white-space:normal">{% for a in linked.get(t['id'], []) %}<a href="{{ url_for('contact_card', site_id=s['id'], contact_id=a['id']) }}">{{ a['name'] }}</a><br>{% else %}{{ t['assignee_text'] or '' }}{% endfor %}</td><td data-label="יעד">{{ t['target_text'] }}</td><td data-label="סטטוס" style="min-width:180px"><span class="pill {{ sclass.get(t['status'], '') }}">{{ t['status'] }}</span>{% if t['id'] in closure_ids %}<br><span class="note">לבדיקה: לא הופיעה בפרוטוקול האחרון</span>{% endif %}{% if t['id'] in last %}<br><span class="note">עודכן ע"י {{ last[t['id']]['changed_by'] }}, {{ last[t['id']]['changed_at'] }}</span>{% endif %}
<form method="post" action="{{ url_for('task_status', task_id=t['id']) }}" data-ajax>
<select name="status" style="width:100%;min-width:150px">{% for o in status_options %}<option>{{ o }}</option>{% endfor %}</select>
<input name="by" placeholder="מי טיפל" style="width:100%;min-width:150px;margin-top:4px"><button style="margin-top:4px">עדכון</button></form></td>
<td data-label="הדפסה"><form method="post" action="{{ url_for('task_print_flag', task_id=t['id']) }}">
<input type="hidden" name="next" value="{{ request.full_path }}">
<input type="checkbox" {% if t['print_flag'] %}checked{% endif %} onchange="this.form.submit()" title="סמן להדפסה"></form></td></tr>{% endfor %}</table>
<h3>לידיעה</h3>
<table class="cards"><tr><th>#</th><th>נושא</th><th>משימה</th><th style="min-width:220px">לטיפול</th></tr>
{% for t in info_t %}<tr><td data-label="#">{{ t['public_no']|pubno }}</td><td data-label="נושא">{{ t['topic'] }}</td><td data-label="משימה" style="width:100%">{{ t['text'] }}</td><td data-label="לטיפול">{% for a in linked.get(t['id'], []) %}<a href="{{ url_for('contact_card', site_id=s['id'], contact_id=a['id']) }}">{{ a['name'] }}</a><br>{% else %}{{ t['assignee_text'] or '' }}{% endfor %}</td></tr>{% else %}<tr><td colspan="4">אין משימות לידיעה</td></tr>{% endfor %}</table>
<h3>משימות שבוצעו</h3>
<table class="cards"><tr><th>#</th><th>נושא</th><th>משימה</th><th style="min-width:220px">לטיפול</th></tr>
{% for t in done_t %}<tr class="done"><td data-label="#">{{ t['public_no']|pubno }}</td><td data-label="נושא">{{ t['topic'] }}</td><td data-label="משימה" style="width:100%">{{ t['text'] }}</td><td data-label="לטיפול">{% for a in linked.get(t['id'], []) %}<a href="{{ url_for('contact_card', site_id=s['id'], contact_id=a['id']) }}">{{ a['name'] }}</a><br>{% else %}{{ t['assignee_text'] or '' }}{% endfor %}</td></tr>{% endfor %}</table>
""", s=s, kp=kp, pct=pct, status_rows=status_rows, site_image=site_image, aged_rows=aged_rows, topic_rows=topic_rows, top_topics=top_topics, is_admin=bool(current_contact() and current_contact()["is_admin"]), sclass=STATUS_CLASS, sort=sort, closure_ids=closure_ids, last=last, status_options=task_status_names(), weekdays=WEEKDAYS, open_t=open_t, info_t=info_t, done_t=done_t, linked=_linked_people(site_id), rowclass=lambda t: {0: "overdue", 1: "urgent", 9: "done"}.get(priority_of(t, today), ""),
        land_txt=_dmy(s['land_date']), delivery_txt=_dmy(s['delivery_date']),
        month_txt=construction_schedule(s["land_date"], s["months_total"])[1],
        end_txt=construction_schedule(s["land_date"], s["months_total"])[0] or "לא נקבע",
        latest=latest, latest_tasks=latest_tasks, latest_date=latest_date, now_txt=dt.datetime.now().strftime("%H:%M:%S"))


PROTOCOL_STEPS = ["העלאת מסמך", "פענוח תוכן", "השוואה להיסטוריה", "בקרת קליטה", "אישור הפיקוח", "עדכון כלל המשתמשים"]
STEPPER_TPL = """<div class="stepper">{% for n in range(1, 7) %}<span class="step {% if n < active %}done{% elif n == active %}active{% endif %}">{{ n }}. {{ step_names[n-1] }}</span>{% endfor %}</div>"""


def delete_site_data(site_id):
    """Remove a site and every record that belongs to it. Child rows (by task) go first,
    while the tasks still exist; then rows that carry site_id; then the site itself."""
    d = db()
    tables = [r[0] for r in d.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
    for t in tables:
        cols = [r[1] for r in d.execute(f"PRAGMA table_info({t})").fetchall()]
        if "task_id" in cols and "site_id" not in cols:
            d.execute(f"DELETE FROM {t} WHERE task_id IN (SELECT id FROM tasks WHERE site_id=?)", (site_id,))
    for t in tables:
        cols = [r[1] for r in d.execute(f"PRAGMA table_info({t})").fetchall()]
        if t != "sites" and "site_id" in cols:
            d.execute(f"DELETE FROM {t} WHERE site_id=?", (site_id,))
    d.execute("DELETE FROM sites WHERE id=?", (site_id,))
    d.commit()


@app.route("/admin")
def admin_overview():
    """General administration: every site, its figures, deletion requests and users with a login."""
    if not current_contact()["is_admin"]:
        abort(403)
    rows = []
    for s_row in db().execute("SELECT * FROM sites ORDER BY id").fetchall():
        kp = site_kpis(s_row["id"])
        total = kp["open"] + kp["closed"]
        meetings = db().execute("SELECT COUNT(*) FROM meetings WHERE site_id=? AND status='published'",
                                (s_row["id"],)).fetchone()[0]
        rows.append({"s": s_row, "kp": kp, "pct": round(100 * kp["closed"] / total) if total else 0, "meetings": meetings})
    users = db().execute("SELECT first_name, last_name, email FROM contacts WHERE password_hash IS NOT NULL ORDER BY email").fetchall()
    return page("ניהול כללי", """
<p><a href="{{ url_for('site_new') }}">+ אתר חדש</a> · <a href="{{ url_for('sites') }}">חזרה לרשימת האתרים</a></p>
<h3>אתרים</h3>
<table class="cards"><tr><th>אתר</th><th>פתוחות</th><th>נסגרו</th><th>אחוז סגירה</th><th>פרוטוקולים שפורסמו</th><th>מחיקה</th></tr>
{% for r in rows %}<tr><td data-label="אתר"><a href="{{ url_for('site_view', site_id=r['s']['id']) }}">{{ r['s']['name'] }}</a></td>
<td data-label="פתוחות">{{ r['kp']['open'] }}</td><td data-label="נסגרו">{{ r['kp']['closed'] }}</td><td data-label="אחוז סגירה">{{ r['pct'] }}%</td><td data-label="פרוטוקולים">{{ r['meetings'] }}</td>
<td data-label="מחיקה"><form method="post" action="{{ url_for('admin_site_delete_request', site_id=r['s']['id']) }}" onsubmit="return confirm('לשלוח בקשת מחיקה לאישור? האתר יימחק רק לאחר אישור בקישור שיישלח לבעל הפרויקט.')"><button>בקשת מחיקה</button></form></td></tr>{% endfor %}</table>
<h3>משתמשים עם גישת מנהל</h3>
<p class="note">{{ users|length }} משתמשים עם חשבון במערכת.</p>
<table class="cards"><tr><th>שם</th><th>דוא"ל</th></tr>{% for u in users %}<tr><td data-label="שם">{{ u['first_name'] or '' }} {{ u['last_name'] or '' }}</td><td data-label='דוא"ל'>{{ u['email'] }}</td></tr>{% endfor %}</table>
""", rows=rows, users=users, url_for=url_for)


@app.route("/admin/sites/<int:site_id>/delete", methods=["POST"])
def admin_site_delete_request(site_id):
    """Deleting a site needs approval: a link goes to the owner's mailbox and nothing is deleted yet."""
    me = current_contact()
    if not me["is_admin"]:
        abort(403)
    s_row = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    token = secrets.token_urlsafe(24)
    now = dt.datetime.now()
    db().execute("INSERT INTO site_delete_requests(token_hash, site_id, requested_by, created_at, expires_at) VALUES (?,?,?,?,?)",
                 (hashlib.sha256(token.encode()).hexdigest(), site_id, me["email"],
                  now.isoformat(timespec="minutes"), (now + dt.timedelta(days=2)).isoformat(timespec="minutes")))
    link = url_for("admin_site_delete_approve", token=token, _external=True)
    body = (f'בקשה למחיקת האתר "{s_row["name"]}" נשלחה על ידי {me["email"]}. '
            f"כל נתוני האתר יימחקו לצמיתות לאחר אישור. לאישור פתחו את הקישור (בתוקף יומיים): {link}")
    enqueue(site_id, dev_contact(), body, "מחיקת אתר", channel="email")
    db().commit()
    return page("בקשת המחיקה נשלחה", '<p>הבקשה נשלחה לאישור בעל הפרויקט. האתר לא יימחק לפני האישור.</p><p><a href="/admin">חזרה לניהול כללי</a></p>')


@app.route("/admin/approve-delete/<token>", methods=["GET", "POST"])
def admin_site_delete_approve(token):
    row = db().execute("SELECT * FROM site_delete_requests WHERE token_hash=?",
                       (hashlib.sha256(token.encode()).hexdigest(),)).fetchone()
    if row is None or row["expires_at"] < dt.datetime.now().isoformat(timespec="minutes"):
        abort(404)
    s_row = db().execute("SELECT * FROM sites WHERE id=?", (row["site_id"],)).fetchone()
    name = s_row["name"] if s_row else ""
    if request.method == "POST":
        if s_row is not None:
            delete_site_data(row["site_id"])
        return page("האתר נמחק", f'<p>האתר "{name}" וכל נתוניו נמחקו.</p>')
    if s_row is None:
        abort(404)
    return page("אישור מחיקת אתר", """
<p>לאישור מחיקת האתר "{{ name }}" וכל נתוניו (משימות, פרוטוקולים, תוכניות ויומנים) לחצו על האישור. הפעולה אינה הפיכה.</p>
<form method="post"><button>אישור מחיקה סופית</button></form>""", name=name)


def reset_site_protocols(site_id):
    """Remove every protocol of a site, its drafts, and the tasks that came from them with
    everything recorded against those tasks. Contacts, the site, plans and manual decisions stay."""
    d = db()
    task_ids = [r[0] for r in d.execute("SELECT id FROM tasks WHERE site_id=?", (site_id,)).fetchall()]
    for table in ("task_status_log", "task_comments", "task_docs", "task_links", "task_assignments",
                  "task_closure_pending", "task_targets", "task_occurrences", "task_images"):
        for tid in task_ids:
            d.execute(f"DELETE FROM {table} WHERE task_id=?", (tid,))
    d.execute("DELETE FROM admin_queue WHERE site_id=?", (site_id,))
    d.execute("DELETE FROM tasks WHERE site_id=?", (site_id,))
    d.execute("DELETE FROM meetings WHERE site_id=?", (site_id,))
    d.commit()


@app.route("/sites/<int:site_id>/guest-link", methods=["POST"])
def site_guest_link(site_id):
    """Administrator only: a link that opens the site read-only, without a password, for one hour."""
    if not current_contact() or not current_contact()["is_admin"]:
        abort(403)
    db().execute("SELECT 1 FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    token = secrets.token_urlsafe(24)
    expires = (dt.datetime.now() + dt.timedelta(hours=1)).isoformat(timespec="minutes")
    email = f"guest-{secrets.token_hex(6)}@guest.invalid"
    cur = db().execute(
        "INSERT INTO contacts(email, first_name, last_name, is_guest, guest_token_hash, expires_at) VALUES (?,?,?,?,?,?)",
        (email, "אורח", "לשעה", 1, hashlib.sha256(token.encode()).hexdigest(), expires))
    db().execute("INSERT INTO memberships(site_id, contact_id, site_role, permission) VALUES (?,?,?,?)",
                 (site_id, cur.lastrowid, "בעלי דירות", "צפייה"))
    db().commit()
    link = url_for("guest_enter", token=token, _external=True)
    return page("קישור גישה לשעה", """<p>הקישור פעיל עד {{ expires }} בלבד, לצפייה בלבד ללא סיסמה.</p>
<p><input value="{{ link }}" readonly style="width:100%" onclick="this.select()"></p>
<p><a href="{{ url_for('site_view', site_id=site_id) }}">חזרה לאתר</a></p>""", expires=expires.replace("T", " "), link=link, site_id=site_id)


@app.route("/guest/<token>")
def guest_enter(token):
    h = hashlib.sha256(token.encode()).hexdigest()
    row = db().execute("SELECT * FROM contacts WHERE is_guest=1 AND guest_token_hash=?", (h,)).fetchone()
    if row is None or not row["expires_at"] or row["expires_at"] < dt.datetime.now().isoformat(timespec="minutes"):
        return page("הקישור פג תוקף", "<p>הקישור אינו פעיל יותר. בקשו מהמנהל קישור חדש.</p>"), 410
    session["cid"] = row["id"]
    m = db().execute("SELECT site_id FROM memberships WHERE contact_id=? LIMIT 1", (row["id"],)).fetchone()
    return redirect(url_for("site_view", site_id=m["site_id"]) if m else url_for("sites"))


@app.route("/sites/<int:site_id>/protocols/reset", methods=["POST"])
def site_protocols_reset(site_id):
    """Administrator only (not listed in SECTION_OF). Needs the typed confirmation."""
    if not current_contact()["is_admin"]:
        abort(403)
    s_row = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    if request.form.get("confirm", "").strip() != "מחק הכול":
        return page("מחיקת פרוטוקולים", """<p>לא נמחק דבר. כדי לאשר יש לכתוב בדיוק: מחק הכול</p>
<p><a href="{{ url_for('site_protocols', site_id=s['id']) }}">חזרה למאגר הפרוטוקולים</a></p>""",
                    s=s_row, url_for=url_for), 400
    reset_site_protocols(site_id)
    return redirect(url_for("site_protocols", site_id=site_id))


@app.route("/sites/<int:site_id>/protocols/publish", methods=["POST"])
def site_protocols_publish(site_id):
    """Publish the chosen drafts (or every draft with all=1), oldest protocol date first."""
    db().execute("SELECT 1 FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    if request.form.get("all"):
        ids = [r[0] for r in db().execute(
            "SELECT id FROM meetings WHERE site_id=? AND status!='published'", (site_id,)).fetchall()]
    else:
        ids = [int(x) for x in request.form.getlist("ids") if x.isdigit()]
    if not ids:
        return redirect(url_for("site_protocols", site_id=site_id))
    marks = ",".join("?" * len(ids))
    drafts = db().execute(
        f"SELECT id FROM meetings WHERE site_id=? AND status!='published' AND id IN ({marks}) "
        "ORDER BY meeting_date IS NULL, meeting_date, id", (site_id, *ids)).fetchall()
    for r in drafts:
        publish_meeting(r["id"])
    return redirect(url_for("site_protocols", site_id=site_id))


@app.route("/sites/<int:site_id>/protocols")
def site_protocols(site_id):
    """Published protocols, newest protocol date first; sort by publish time or filter to this month."""
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    sort = request.args.get("sort", "date")
    this_month = request.args.get("month") == "current"
    month_key = dt.date.today().strftime("%Y-%m")
    on_date = request.args.get("on", "")  # exact protocol date, ISO
    ym = request.args.get("ym", "") if not this_month else month_key  # a month, YYYY-MM
    items = []
    for m in db().execute("SELECT * FROM meetings WHERE site_id=? AND status='published'", (site_id,)).fetchall():
        if this_month and not (m["meeting_date"] or "").startswith(month_key):
            continue
        if on_date and m["meeting_date"] != on_date:
            continue
        if ym and not (m["meeting_date"] or "").startswith(ym):
            continue
        rows = [r for r in json.loads(m["staged_rows"] or "[]") if not r.get("skip")]
        items.append({"m": m, "rows": len(rows)})
    if sort == "published":
        items.sort(key=lambda i: (i["m"]["published_at"] or "", i["m"]["id"]), reverse=True)
    else:
        items.sort(key=lambda i: (i["m"]["meeting_date"] or "", i["m"]["published_at"] or ""), reverse=True)
    drafts = []
    for m in db().execute("SELECT * FROM meetings WHERE site_id=? AND status!='published' ORDER BY id DESC", (site_id,)).fetchall():
        drafts.append({"m": m, "rows": len([r for r in json.loads(m["staged_rows"] or "[]") if not r.get("skip")])})
    return page("מאגר פרוטוקולים - " + s["name"], """
<p><a href="{{ url_for('site_view', site_id=s['id']) }}">חזרה לאתר</a></p>
<p>
<a href="{{ url_for('site_protocols', site_id=s['id'], sort=sort) }}" {% if not this_month %}style="font-weight:bold"{% endif %}>כל הפרוטוקולים</a> · <a href="{{ url_for('site_images', site_id=s['id']) }}">תמונות לפי נושא</a> ·
<a href="{{ url_for('site_protocols', site_id=s['id'], sort=sort, month='current') }}" {% if this_month %}style="font-weight:bold"{% endif %}>החודש</a>
</p>
<form method="get" style="margin-bottom:8px">{% if this_month %}<input type="hidden" name="month" value="current">{% endif %}
מיון לפי <select name="sort">
<option value="date" {% if sort != 'published' %}selected{% endif %}>תאריך הפרוטוקול (החדשים קודם)</option>
<option value="published" {% if sort == 'published' %}selected{% endif %}>זמן הפרסום (החדשים קודם)</option>
</select> <button>מיון</button></form>
<form method="get" style="margin-bottom:8px">תאריך מסוים <input type="date" name="on" value="{{ on_date }}"> חודש <input type="month" name="ym" value="{{ ym }}"> <button>הצג</button> <a href="{{ url_for('site_protocols', site_id=s['id']) }}">הצג הכול</a></form>
<table class="cards"><tr><th>סוג</th><th>תאריך הפרוטוקול</th><th>זמן הפרסום</th><th>קובץ</th><th>שורות</th><th></th></tr>
{% for i in items %}<tr><td data-label="סוג">{{ i['m']['kind'] }}</td><td data-label="תאריך הפרוטוקול">{{ i['m']['meeting_date'] or 'לא זוהה' }}</td>
<td data-label="זמן הפרסום">{{ i['m']['published_at'] or '' }}</td><td data-label="קובץ">{{ i['m']['source_name'] or '' }}</td><td data-label="שורות">{{ i['rows'] }}</td>
<td data-label=""><a href="{{ url_for('meeting_view', meeting_id=i['m']['id']) }}">פתיחה</a></td></tr>
{% else %}<tr><td colspan="6">{{ 'אין פרוטוקולים מפורסמים בחודש הנוכחי' if this_month else 'אין פרוטוקולים מפורסמים' }}</td></tr>{% endfor %}</table>
<h3>הועלו ולא פורסמו (טיוטות)</h3>
<form method="post" action="{{ url_for('site_protocols_publish', site_id=s['id']) }}">
<p><button>פרסם את הנבחרים</button> <button name="all" value="1" onclick="return confirm('לפרסם את כל הטיוטות לפי סדר התאריכים?')">פרסם הכל</button> <span class="note">מסומנים: הפרוטוקולים יפורסמו מהישן לחדש.</span></p>
<table class="cards"><tr><th></th><th>סוג</th><th>תאריך הפרוטוקול</th><th>קובץ</th><th>שורות שזוהו</th><th></th></tr>
{% for d in drafts %}<tr><td data-label="בחירה"><input type="checkbox" name="ids" value="{{ d['m']['id'] }}"></td><td data-label="סוג">{{ d['m']['kind'] }}</td><td data-label="תאריך הפרוטוקול">{{ d['m']['meeting_date'] or 'לא זוהה' }}</td>
<td data-label="קובץ">{{ d['m']['source_name'] or '' }}</td><td data-label="שורות שזוהו">{{ d['rows'] }}{% if d['rows'] == 0 %} <span class="note">לא זוהו שורות - בדקו את הקובץ</span>{% endif %}</td>
<td data-label=""><a href="{{ url_for('meeting_review', meeting_id=d['m']['id']) }}">בדיקה ופרסום</a></td></tr>
{% else %}<tr><td colspan="6">אין טיוטות</td></tr>{% endfor %}</table></form>
{% if is_admin %}<h3>מחיקת כל הפרוטוקולים (למנהל בלבד)</h3>
<form method="post" action="{{ url_for('site_protocols_reset', site_id=s['id']) }}" onsubmit="return confirm('למחוק את כל הפרוטוקולים והמשימות של האתר? הפעולה אינה הפיכה.')">
<p class="note">ימחקו כל הפרוטוקולים והטיוטות, המשימות שנוצרו מהם, ההיסטוריה, השיוכים, ההערות והקבצים המקושרים למשימות. אנשי הקשר, האתר, התוכניות והחלטות שנרשמו ידנית נשארים.</p>
<input name="confirm" placeholder='כתבו "מחק הכול"'> <button>מחיקה</button></form>{% endif %}
""", s=s, items=items, sort=sort, this_month=this_month, drafts=drafts, on_date=on_date, ym=ym, is_admin=current_contact()["is_admin"])


FILE_DATE = re.compile(r"(?<!\d)(\d{2})\.(\d{2})\.(\d{4}|\d{2})(?!\d)")


def date_from_filename(name):
    """A date written in a file name (dd.mm.yy or dd.mm.yyyy) as an ISO date, or None."""
    m = FILE_DATE.search(name or "")
    if not m:
        return None
    year = int(m.group(3)) + (2000 if len(m.group(3)) == 2 else 0)
    try:
        return dt.date(year, int(m.group(2)), int(m.group(1))).isoformat()
    except ValueError:
        return None


@app.route("/sites/<int:site_id>/meetings/new", methods=["GET", "POST"])
def meeting_upload(site_id):
    if request.method == "POST":
        kind = request.form.get("kind", "ביצוע")
        ids = []
        for f in request.files.getlist("file"):
            if not f or not f.filename:
                continue
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            f.save(tmp.name)
            rows = extract_rows(tmp.name, image_dir=os.path.join(UPLOAD_DIR, "protocol_images", str(site_id)))
            # a page with no date of its own takes the date written in the file name
            named = date_from_filename(f.filename)
            for r in rows:
                if not r.get("meeting_date") and named:
                    r["meeting_date"] = dt.date.fromisoformat(named).strftime("%d.%m.%y")
            for day, group in split_rows_by_date(rows):
                cur = db().execute(
                    "INSERT INTO meetings(site_id, kind, source_name, staged_rows, meeting_date) VALUES (?,?,?,?,?)",
                    (site_id, kind, f.filename, json.dumps(group, ensure_ascii=False), day))
                ids.append(cur.lastrowid)
        db().commit()
        if not ids:
            return redirect(url_for("meeting_upload", site_id=site_id))
        if len(ids) == 1:
            return redirect(url_for("meeting_review", meeting_id=ids[0]))
        # several protocols: they wait as drafts in the protocol repository, filed by date
        return redirect(url_for("site_protocols", site_id=site_id))
    return page("העלאת פרוטוקולים", STEPPER_TPL + """
<form method="post" enctype="multipart/form-data">
<p>סוג פגישה <select name="kind"><option>ביצוע</option><option>שבועית</option><option>סופרפוזיציה</option></select></p>
<p>אפשר לבחור כמה קבצים בבת אחת. תאריך הפרוטוקול נקבע לפי התאריך שבקובץ, ואם אין כזה - לפי התאריך שבשם הקובץ.</p>
<input type="file" name="file" accept=".pdf" multiple required> <button>העלאה לבדיקה</button></form>""", url_for=url_for, active=1, step_names=PROTOCOL_STEPS)


@app.route("/protocol-images/<int:site_id>/<name>")
def protocol_image(site_id, name):
    if not re.fullmatch(r"[0-9a-f]{16}\.png", name):
        abort(404)
    return send_from_directory(os.path.join(UPLOAD_DIR, "protocol_images", str(site_id)), name)


@app.route("/sites/<int:site_id>/images")
def site_images(site_id):
    """Pictures from the protocols, filed by topic. ?topic= shows one topic."""
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    topic = request.args.get("topic", "")
    rows = db().execute(
        """SELECT i.file, i.topic, i.task_id, i.meeting_id, t.public_no, t.text
           FROM task_images i JOIN tasks t ON t.id = i.task_id
           WHERE i.site_id=? ORDER BY i.topic, i.id""", (site_id,)).fetchall()
    topics = sorted({r["topic"] for r in rows})
    shown = [r for r in rows if not topic or r["topic"] == topic]
    return page("תמונות מהפרוטוקולים - " + s["name"], """
<p><a href="{{ url_for('site_view', site_id=s['id']) }}">חזרה לאתר</a></p>
<p>נושאים: <a href="{{ url_for('site_images', site_id=s['id']) }}">הכול</a>
{% for t in topics %} · <a href="{{ url_for('site_images', site_id=s['id'], topic=t) }}">{{ t }}</a>{% endfor %}</p>
<div style="display:flex;flex-wrap:wrap;gap:10px">
{% for r in shown %}<figure style="margin:0;width:200px">
<a href="{{ url_for('protocol_image', site_id=s['id'], name=r['file']) }}" target="_blank"><img src="{{ url_for('protocol_image', site_id=s['id'], name=r['file']) }}" style="width:200px;border:1px solid #ddd"></a>
<figcaption style="font-size:12px">{{ r['topic'] }} · {{ r['public_no']|pubno }}<br><a href="{{ url_for('meeting_view', meeting_id=r['meeting_id']) }}">הפרוטוקול</a></figcaption></figure>
{% else %}<p>אין תמונות{% if topic %} בנושא זה{% endif %}.</p>{% endfor %}
</div>""", s=s, topics=topics, shown=shown)


@app.route("/meetings/<int:meeting_id>/view")
def meeting_view(meeting_id):
    """A published protocol as it was written: its own rows, with the status each task has today."""
    m = db().execute("SELECT * FROM meetings WHERE id=?", (meeting_id,)).fetchone() or abort(404)
    rows = db().execute(
        """SELECT o.number, o.topic, o.text, o.target, t.status, t.id AS task_id
           FROM task_occurrences o JOIN tasks t ON t.id = o.task_id WHERE o.meeting_id=? ORDER BY o.id""",
        (meeting_id,)).fetchall()
    try:
        date_txt = dt.date.fromisoformat(m["meeting_date"]).strftime("%d.%m.%Y")
    except (TypeError, ValueError):
        date_txt = "לא זוהה"
    pics = {}
    for pr in db().execute("SELECT task_id, file FROM task_images WHERE meeting_id=? ORDER BY id", (meeting_id,)).fetchall():
        pics.setdefault(pr["task_id"], []).append(pr["file"])
    return page(f"פרוטוקול {date_txt}", """
<p><a href="{{ url_for('site_protocols', site_id=m['site_id']) }}">חזרה למאגר הפרוטוקולים</a></p>
<p>סוג: {{ m['kind'] }} · תאריך: {{ date_txt }} · קובץ: {{ m['source_name'] or '' }}</p>
<table class="cards"><tr><th>#</th><th>נושא</th><th>משימה</th><th>יעד</th><th>סטטוס היום</th></tr>
{% for r in rows %}<tr><td data-label="#">{{ r['number'] }}</td><td data-label="נושא">{{ r['topic'] or '' }}</td>
<td data-label="משימה" style="width:100%">{{ r['text'] }}{% for f in pics.get(r['task_id'], []) %} <a href="{{ url_for('protocol_image', site_id=m['site_id'], name=f) }}" target="_blank"><img src="{{ url_for('protocol_image', site_id=m['site_id'], name=f) }}" style="max-width:160px;max-height:110px;display:block;margin-top:4px;border:1px solid #ddd"></a>{% endfor %}</td><td data-label="יעד">{{ r['target'] or '' }}</td>
<td data-label="סטטוס היום"><span class="pill {{ sclass.get(r['status'], '') }}">{{ r['status'] }}</span>
<form method="post" action="{{ url_for('task_status', task_id=r['task_id']) }}" data-ajax>
<select name="status" style="width:100%;min-width:150px">{% for o in options %}<option {% if o == r['status'] %}selected{% endif %}>{{ o }}</option>{% endfor %}</select>
<input name="by" placeholder="מי טיפל" style="width:100%;min-width:150px;margin-top:4px"><button style="margin-top:4px">עדכון</button></form></td></tr>
{% else %}<tr><td colspan="5">אין שורות בפרוטוקול זה</td></tr>{% endfor %}</table>""",
        m=m, rows=rows, date_txt=date_txt, sclass=STATUS_CLASS, options=task_status_names(), pics=pics)


@app.route("/meetings/<int:meeting_id>/review", methods=["GET", "POST"])
def meeting_review(meeting_id):
    m = db().execute("SELECT * FROM meetings WHERE id=?", (meeting_id,)).fetchone() or abort(404)
    rows = json.loads(m["staged_rows"] or "[]")
    if request.method == "POST":
        for i, r in enumerate(rows):
            r["number"] = request.form.get(f"number_{i}", r["number"])
            r["topic"] = request.form.get(f"topic_{i}", r["topic"])
            r["task"] = request.form.get(f"task_{i}", r["task"])
            r["assignee"] = request.form.get(f"assignee_{i}", r["assignee"])
            r["target"] = request.form.get(f"target_{i}", r["target"])
            r["skip"] = bool(request.form.get(f"skip_{i}"))
        db().execute("UPDATE meetings SET staged_rows=? WHERE id=?",
                     (json.dumps(rows, ensure_ascii=False), meeting_id))
        if request.form.get("action") == "publish":
            publish_meeting(meeting_id)
            return redirect(url_for("site_view", site_id=m["site_id"]))
        db().commit()
        flash("נשמר כטיוטה")
    return page("בדיקת פרוטוקול לפני פרסום", STEPPER_TPL + """
<p>ערכו כל שורה לפני הפרסום. סמנו 'לא לפרסם' לשורות שחולצו בצורה שגויה.</p>
<form method="post"><table><tr><th>#</th><th>נושא</th><th>משימה</th><th>לטיפול</th><th>יעד</th><th>לא לפרסם</th></tr>
{% for r in rows %}{% set i = loop.index0 %}<tr>
<td><input name="number_{{i}}" value="{{ r['number'] }}" size="5"></td>
<td><input name="topic_{{i}}" value="{{ r['topic'] }}"></td>
<td><textarea name="task_{{i}}" rows="3" cols="50">{{ r['task'] }}</textarea></td>
<td><input name="assignee_{{i}}" value="{{ r['assignee'] }}"></td>
<td><input name="target_{{i}}" value="{{ r['target'] }}"></td>
<td><input type="checkbox" name="skip_{{i}}" {% if r.get('skip') %}checked{% endif %}></td></tr>{% endfor %}</table>
<p><button name="action" value="save">שמירה כטיוטה</button> <button name="action" value="publish">אישור ופרסום</button></p></form>
""", rows=rows, active=4, step_names=PROTOCOL_STEPS)


# Render (gunicorn) imports this module and never runs the __main__ block,
# so the database schema must be created here on the server.
if os.environ.get("RENDER"):
    init_db()


if __name__ == "__main__":
    import sys
    if len(sys.argv) == 4 and sys.argv[1] == "set-password":
        set_password_cli(sys.argv[2], sys.argv[3])
    elif len(sys.argv) == 2 and sys.argv[1] == "run-reminders":
        init_db()
        with app.app_context():
            for site in db().execute("SELECT id FROM sites").fetchall():
                send_task_reminders(site["id"])
            db().commit()
            print("queued:", db().execute("SELECT COUNT(*) FROM outbox WHERE status='queued'").fetchone()[0])
    else:
        init_db()
        app.run(debug=False, port=5055)
