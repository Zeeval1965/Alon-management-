"""MVP skeleton: sites, contacts, protocol import with approval, tasks, outbox."""
import datetime as dt
import json
import os
import re
import sqlite3
import tempfile
import uuid

import openpyxl
from flask import Flask, abort, flash, g, redirect, render_template_string, request, send_from_directory, session, url_for
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
CREATE TABLE IF NOT EXISTS plans (
  id INTEGER PRIMARY KEY, site_id INTEGER NOT NULL, title TEXT NOT NULL, version TEXT,
  file_path TEXT NOT NULL, original_name TEXT, uploaded_at TEXT
);
CREATE TABLE IF NOT EXISTS stage_status (
  site_id INTEGER NOT NULL, level TEXT NOT NULL, stage TEXT NOT NULL, status TEXT NOT NULL,
  PRIMARY KEY(site_id, level, stage)
);
"""

app = Flask(__name__)
app.secret_key = os.environ.get("APP_SECRET", "dev-only")


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


def init_db():
    c = sqlite3.connect(DB_PATH)
    c.executescript(SCHEMA)
    for stmt in ("ALTER TABLE task_assignments ADD COLUMN is_responsible INTEGER DEFAULT 0",
                 "ALTER TABLE task_assignments ADD COLUMN is_cc INTEGER DEFAULT 0",
                 "ALTER TABLE memberships ADD COLUMN notify_freq TEXT",
                 "ALTER TABLE sites ADD COLUMN digest_time TEXT DEFAULT '08:00'",
                 "ALTER TABLE sites ADD COLUMN weekly_day TEXT DEFAULT 'ראשון'",
                 "ALTER TABLE outbox ADD COLUMN mode TEXT DEFAULT 'מיידי'",
                 "ALTER TABLE contacts ADD COLUMN password_hash TEXT",
                 "ALTER TABLE contacts ADD COLUMN is_admin INTEGER DEFAULT 0",
                 "ALTER TABLE sites ADD COLUMN levels TEXT",
                 "ALTER TABLE sites ADD COLUMN stages TEXT"):
        try:
            c.execute(stmt)
        except sqlite3.OperationalError:
            pass  # column already exists
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
                       target_text, target_date, last_meeting_id)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                (site_id, public, r["number"], r["topic"], r["task"], r["assignee"],
                 r["target"], target_date, meeting_id))
            task_id = cur.lastrowid
        else:
            task_id = match["id"]
            if match["target_text"] != r["target"]:  # new estimate from this protocol
                db().execute(
                    "UPDATE tasks SET target_text=?, target_date=?, last_meeting_id=? WHERE id=?",
                    (r["target"], target_date, meeting_id, task_id))
        db().execute("INSERT INTO task_targets(task_id, target_text, meeting_id) VALUES (?,?,?)",
                     (task_id, r["target"], meeting_id))
        if re.search(r"(^|\s)בוצע$", (r.get("target") or "").strip()):  # 'בוצע' or 'dd.mm.yy בוצע'
            db().execute("UPDATE tasks SET status='בוצע' WHERE id=?", (task_id,))
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
               WHERE t.site_id=? AND t.status!='בוצע' AND lm.kind=? AND t.last_meeting_id != ?""",
            (site_id, m["kind"], meeting_id)).fetchall()
        for t in gone:
            if t["id"] not in listed_ids:
                db().execute("UPDATE tasks SET status='בוצע' WHERE id=?", (t["id"],))
    for task_id in published:
        suggest_assignees(task_id)
    db().execute("UPDATE meetings SET status='published' WHERE id=?", (meeting_id,))
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


def enqueue(site_id, contact, body, mode):
    if DEV_MODE:
        intended = f"{contact['first_name']} {contact['last_name']}".strip() or contact["email"]
        body = f"[פיתוח - במקור אל: {intended}] {body}"
        contact = dev_contact()
    if DEV_MODE:
        channel = "email"  # until WhatsApp is connected, development messages go to the owner's email
    else:
        channel = "email" if contact["is_consultant"] else (contact["pref_channel"] or "whatsapp")
    db().execute(
        "INSERT INTO outbox(site_id, contact_id, channel, body, mode, created) VALUES (?,?,?,?,?,?)",
        (site_id, contact["id"], channel, body, mode, dt.datetime.now().isoformat(timespec="minutes")))


def notify_task_event(task, text):
    for contact, mode in recipients_for_event(task["site_id"]):
        enqueue(task["site_id"], contact, text, mode)
    db().commit()


def notify_supervisors_task_handled(task, by_name, new_status):
    text = (f"משימה {task['public_no']} ({task['topic'] or 'ללא נושא'}): {task['text'][:120]} "
            f"- עודכנה לסטטוס '{new_status}' על ידי {by_name}.")
    notify_task_event(task, text)


def send_task_reminders(site_id, today=None):
    """Queue reminders for open tasks that are overdue or due within a day (site default time)."""
    today = today or dt.date.today()
    limit = (today + dt.timedelta(days=1)).isoformat()
    tasks = db().execute(
        """SELECT * FROM tasks WHERE site_id=? AND status NOT IN ('בוצע', 'לידיעה')
           AND (target_date IS NOT NULL AND target_date <= ?)""", (site_id, limit)).fetchall()
    for t in tasks:
        when = "באיחור" if t["target_date"] < today.isoformat() else "תאריך היעד הוא מחר או היום"
        text = (f"תזכורת: משימה {t['public_no']} ({t['topic'] or 'ללא נושא'}): {t['text'][:120]} - {when}.")
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
    "site_view": "tasks", "task_status": "tasks", "assignments": "tasks", "confirm_assignee": "tasks",
    "set_responsible": "tasks", "add_responsible": "tasks", "toggle_cc": "tasks", "site_contacts": "tasks", "cc_report": "tasks",
    "site_warnings": "tasks", "assignments_rematch": "tasks", "contact_card": "tasks", "task_triage": "tasks",
    "assign_task": "tasks",
    "decisions_view": "decisions",
    "supervision_view": "supervision",
    "plans_view": "plans", "plan_file": "plans",
    "meeting_upload": "protocols", "meeting_review": "protocols",
    "progress_view": "progress", "progress_save": "progress",
}
PUBLIC_ENDPOINTS = {"login", "logout", "static"}
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
    return g._me


def roles_on_site(contact, site_id):
    rows = db().execute("SELECT site_role FROM memberships WHERE site_id=? AND contact_id=?",
                        (site_id, contact["id"])).fetchall()
    return {r["site_role"] for r in rows}


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
                   target_text, target_date, last_meeting_id) VALUES (?,?,?,?,?,?,?,?,?)""",
            (q["site_id"], public, r["number"], r["topic"], r["task"], r["assignee"],
             r["target"], current_target_date(r["target"]), q["meeting_id"]))
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
    if not db().execute("SELECT 1 FROM task_assignments WHERE task_id=? AND is_responsible=1", (task_id,)).fetchone():
        db().execute("UPDATE task_assignments SET is_responsible=1 WHERE task_id=? AND contact_id=?", (task_id, cid))
    db().commit()
    return redirect(url_for("site_warnings", site_id=t["site_id"]))


@app.route("/tasks/<int:task_id>/triage", methods=["POST"])
def task_triage(task_id):
    """For a task with no one to assign: 'info' marks it for information only, 'delete' removes it."""
    t = db().execute("SELECT site_id FROM tasks WHERE id=?", (task_id,)).fetchone() or abort(404)
    action = request.form.get("action")
    if action == "info":
        db().execute("UPDATE tasks SET status='לידיעה' WHERE id=?", (task_id,))
    elif action == "delete":
        db().execute("DELETE FROM task_assignments WHERE task_id=?", (task_id,))
        db().execute("DELETE FROM task_targets WHERE task_id=?", (task_id,))
        db().execute("DELETE FROM tasks WHERE id=?", (task_id,))
    else:
        abort(400)
    db().commit()
    return redirect(url_for("site_warnings", site_id=t["site_id"]))


@app.route("/tasks/<int:task_id>/status", methods=["POST"])
def task_status(task_id):
    t = db().execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone() or abort(404)
    new_status = request.form["status"]
    by_name = request.form.get("by", "לא צוין")
    db().execute("UPDATE tasks SET status=? WHERE id=?", (new_status, task_id))
    db().commit()
    if new_status != "פתוחה":
        notify_supervisors_task_handled(t, by_name, new_status)
    return redirect(url_for("site_view", site_id=t["site_id"]))


# ---------- assignee linking: text in 'לטיפול' -> contacts (supervisor confirms) ----------
def _full_names(site_id):
    rows = db().execute(
        """SELECT DISTINCT c.id, c.first_name, c.last_name FROM memberships m
           JOIN contacts c ON c.id = m.contact_id WHERE m.site_id=?""", (site_id,)).fetchall()
    return rows


def rematch_site(site_id):
    """Run name matching again for the site's open tasks (after contacts were imported or changed)."""
    for t in db().execute("SELECT id FROM tasks WHERE site_id=? AND status!='בוצע'", (site_id,)).fetchall():
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
    for cid, raw, full in found:
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
        people[key]["tasks"].append(f"{r['public_no']} ({r['topic'] or 'ללא נושא'})")
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
{% for r in rows %}<tr><td>{{ r['public_no'] }}</td><td style="min-width:260px">{{ r['topic'] or '' }}: {{ r['text'] }}</td><td>{{ r['assignee_text'] }}</td><td>{{ r['raw_name'] }}</td>
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
           WHERE t.site_id=? AND t.status NOT IN ('בוצע', 'לידיעה')
             AND NOT EXISTS (SELECT 1 FROM task_assignments a WHERE a.task_id=t.id)
           ORDER BY t.public_no""", (site_id,)).fetchall()
    pending = db().execute(
        """SELECT t.id AS task_id, t.public_no, t.topic, t.text, a.raw_name, a.contact_id, c.first_name, c.last_name, c.org
           FROM task_assignments a JOIN tasks t ON t.id = a.task_id JOIN contacts c ON c.id = a.contact_id
           WHERE t.site_id=? AND t.status!='בוצע' AND a.confirmed=0 ORDER BY t.public_no""", (site_id,)).fetchall()
    choices = db().execute(
        """SELECT DISTINCT c.id, c.first_name, c.last_name FROM memberships m JOIN contacts c ON c.id = m.contact_id
           WHERE m.site_id=? ORDER BY c.first_name, c.last_name""", (site_id,)).fetchall()
    return page("התרעות - " + s["name"], """
<h3>אנשי קשר כפולים (אותו שם, יותר מרשומה אחת) - מומלץ למחוק את הכפילות</h3>
<table><tr><th>שם</th><th>רשומות</th></tr>
{% for d in duplicates %}<tr><td>{{ d['name'] }}</td><td>{% for r in d['records'] %}{{ r['email'] }} ({{ r['org'] }}: {{ r['roles']|join(', ') }})<br>{% endfor %}</td></tr>
{% else %}<tr><td colspan="2">אין כפילויות</td></tr>{% endfor %}</table>
<h3>שיוכים הממתינים לאישור המפקח (שם חלקי או שם משותף לכמה אנשים)</h3>
<table><tr><th>משימה</th><th>תיאור המשימה</th><th>הטקסט בפרוטוקול</th><th>איש קשר מוצע</th><th>ארגון</th></tr>
{% for p in pending %}<tr><td>{{ p['public_no'] }}</td><td style="min-width:260px">{{ p['topic'] or '' }}: {{ p['text'] }}</td>
<td>{{ p['raw_name'] }}</td>
<td><form method="post" action="{{ url_for('assign_task', task_id=p['task_id']) }}">
<select name="contact_id">{% for c in choices %}<option value="{{ c['id'] }}" {% if c['id'] == p['contact_id'] %}selected{% endif %}>{{ c['first_name'] }} {{ c['last_name'] or '' }}</option>{% endfor %}</select>
<button>שייך</button></form></td><td>{{ p['org'] or '' }}</td></tr>
{% else %}<tr><td colspan="5">אין שיוכים ממתינים. לאישור השיוכים: <a href="{{ url_for('assignments', site_id=s['id']) }}">שיוך משימות</a></td></tr>{% endfor %}</table>
<h3>משימות פתוחות ללא שיוך לאיש קשר</h3>
<table><tr><th>#</th><th>נושא</th><th>משימה</th><th>לטיפול (כפי שכתוב)</th></tr>
{% for t in unassigned %}<tr><td>{{ t['public_no'] }}</td><td>{{ t['topic'] or '' }}</td><td>{{ t['text'] }}</td>
<td>{{ t['assignee_text'] or '' }}
<form method="post" action="{{ url_for('assign_task', task_id=t['id']) }}" style="margin-top:4px">
<select name="contact_id">{% for c in choices %}<option value="{{ c['id'] }}">{{ c['first_name'] }} {{ c['last_name'] or '' }}</option>{% endfor %}</select>
<button>שייך</button></form>
<form method="post" action="{{ url_for('task_triage', task_id=t['id']) }}" style="margin-top:4px">
<button name="action" value="info">1 - לידיעה</button>
<button name="action" value="delete" onclick="return confirm('למחוק את המשימה? הפעולה אינה הפיכה.')">2 - מחק</button></form></td></tr>
{% else %}<tr><td colspan="4">כל המשימות הפתוחות משויכות</td></tr>{% endfor %}</table>
<p><a href="{{ url_for('site_view', site_id=s['id']) }}">חזרה לאתר</a></p>
""", duplicates=duplicates, unassigned=unassigned, pending=pending, choices=choices, s=s, url_for=url_for)


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
<tr><th>דוא"ל</th><td dir="ltr">{{ c['email'] or '' }}</td></tr>
<tr><th>כתובת</th><td>{{ c['address'] or '' }}</td></tr>
</table>
<h3>משימות</h3>
<table><tr><th>משימה</th><th>סטטוס</th><th>תפקיד בשיוך</th></tr>
{% for t in tasks %}<tr><td>{{ t['public_no'] }} {{ t['topic'] or '' }}</td><td>{{ t['status'] }}</td>
<td>{% if not t['confirmed'] %}ממתין לאישור{% elif t['is_responsible'] %}אחראי לביצוע{% elif t['is_cc'] %}מכותב{% else %}לצפייה{% endif %}</td></tr>
{% else %}<tr><td colspan="3">אין משימות משויכות</td></tr>{% endfor %}</table>
<p><a href="{{ url_for('site_contacts', site_id=s['id']) }}">חזרה לרשימת אנשי הקשר</a></p>
""", c=c, roles=roles, tasks=tasks, s=s, url_for=url_for)


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
<th>שם</th><th>ארגון</th><th>תפקיד באתר</th><th>טלפון</th><th>נייד</th><th>דוא"ל</th>{% if can_edit %}<th></th>{% endif %}</tr>
{% for r in rows %}<tr>{% if is_admin %}<td><input type="checkbox" class="pick" name="ids" value="{{ r['id'] }}"></td>{% endif %}
<td>{{ r['first_name'] }} {{ r['last_name'] or '' }}</td><td>{{ r['org'] or '' }}</td>
<td>{{ r['site_roles'] }}</td><td>{{ r['phone'] or '' }}</td><td>{{ r['mobile'] or '' }}</td><td>{{ r['email'] or '' }}</td>
{% if can_edit %}<td><a href="{{ url_for('contact_edit', contact_id=r['id']) }}">עריכה</a></td>{% endif %}</tr>{% endfor %}</table>
{% if is_admin %}<input type="hidden" name="site_id" value="{{ s['id'] }}"><p><button>מחיקת הנבחרים</button> <span class="note">המחיקה מוחקת את האדם מכל האתרים. מנהל המערכת לא נמחק.</span></p>{% endif %}
{% if is_admin %}</form>{% endif %}
<p><a href="{{ url_for('site_view', site_id=s['id']) }}">חזרה לאתר</a></p>
""", rows=rows, s=s, url_for=url_for, can_edit=is_manager(current_contact()),
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
</style></head><body>
<header><b>הלל הזקן 14</b>
<a href="{{ url_for('sites') }}">אתרים</a><a href="{{ url_for('contacts_import') }}">ייבוא אנשי קשר</a></header>
<main><h2>{{ title }}</h2>
{% for m in get_flashed_messages() %}<p class="flash">{{ m }}</p>{% endfor %}
{{ body|safe }}</main></body></html>"""


def page(title, body, **ctx):
    return render_template_string(PAGE, title=title, body=render_template_string(body, **ctx),
                                  url_for=url_for, get_flashed_messages=lambda: [])


@app.route("/")
def sites():
    rows = db().execute("SELECT * FROM sites ORDER BY id").fetchall()
    return page("אתרים", """
<p><a href="{{ url_for('site_new') }}">אתר חדש</a></p>
<table><tr><th>אתר</th><th>עלייה לקרקע</th><th>מסירה</th><th>חודשי ביצוע (חוזה)</th><th></th></tr>
{% for s in rows %}<tr><td>{{ s['name'] }}</td><td>{{ s['land_date'] }}</td><td>{{ s['delivery_date'] }}</td>
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


@app.route("/sites/<int:site_id>")
def site_view(site_id):
    s = db().execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone() or abort(404)
    tasks = db().execute("SELECT * FROM tasks WHERE site_id=?", (site_id,)).fetchall()
    today = dt.date.today().isoformat()
    open_t = sorted([t for t in tasks if t["status"] != "בוצע"], key=lambda t: (priority_of(t, today), t["public_no"]))
    done_t = [t for t in tasks if t["status"] == "בוצע"]
    month = db().execute("SELECT month_current, month_total, meeting_date FROM meetings WHERE site_id=? AND status='published' AND month_current IS NOT NULL ORDER BY id DESC LIMIT 1", (site_id,)).fetchone()
    return page(s["name"], """
<form method="post" action="{{ url_for('site_schedule', site_id=s['id']) }}">
שעת שליחה לסיכומים (ברירת מחדל לכל האתר) <input name="digest_time" value="{{ s['digest_time'] }}" size="5">
יום לסיכום שבועי <select name="weekly_day">{% for d in weekdays %}<option {% if d == s['weekly_day'] %}selected{% endif %}>{{ d }}</option>{% endfor %}</select>
<button>שמירה</button></form>
<p>עלייה לקרקע {{ s['land_date'] }} · מסירה לפי חוזה {{ s['delivery_date'] }} · חודש ביצוע {{ month_txt }}</p>
<p><a href="{{ url_for('meeting_upload', site_id=s['id']) }}">העלאת פרוטוקול</a> · <a href="{{ url_for('assignments', site_id=s['id']) }}">שיוך משימות לאנשי קשר</a> · <a href="{{ url_for('site_contacts', site_id=s['id']) }}">אנשי קשר של הפרויקט</a> · <a href="{{ url_for('site_warnings', site_id=s['id']) }}">התרעות לתיקון</a></p>
<h3>משימות פתוחות</h3>
<table><tr><th>#</th><th>נושא</th><th>משימה</th><th style="min-width:220px">לטיפול</th><th>יעד</th><th>סטטוס</th></tr>
{% for t in open_t %}<tr class="{{ rowclass(t) }}"><td>{{ t['public_no'] }}</td><td>{{ t['topic'] }}</td>
<td>{{ t['text'] }}</td><td style="min-width:220px;white-space:normal">{% for a in linked.get(t['id'], []) %}<a href="{{ url_for('contact_card', site_id=s['id'], contact_id=a['id']) }}">{{ a['name'] }}</a>{% if not loop.last %}, {% endif %}{% else %}{{ t['assignee_text'] or '' }}{% endfor %}</td><td>{{ t['target_text'] }}</td><td>{{ t['status'] }}
<form method="post" action="{{ url_for('task_status', task_id=t['id']) }}">
<select name="status"><option>בטיפול</option><option>ממתינה לאישור</option><option>לידיעה</option><option>בוצע</option></select>
<input name="by" placeholder="מי טיפל" size="10"><button>עדכון</button></form></td></tr>{% endfor %}</table>
<h3>משימות שבוצעו</h3>
<table><tr><th>#</th><th>נושא</th><th>משימה</th><th style="min-width:220px">לטיפול</th></tr>
{% for t in done_t %}<tr class="done"><td>{{ t['public_no'] }}</td><td>{{ t['topic'] }}</td><td>{{ t['text'] }}</td><td>{% for a in linked.get(t['id'], []) %}<a href="{{ url_for('contact_card', site_id=s['id'], contact_id=a['id']) }}">{{ a['name'] }}</a>{% if not loop.last %}, {% endif %}{% else %}{{ t['assignee_text'] or '' }}{% endfor %}</td></tr>{% endfor %}</table>
""", s=s, weekdays=WEEKDAYS, open_t=open_t, done_t=done_t, linked=_linked_people(site_id), rowclass=lambda t: {0: "overdue", 1: "urgent", 9: "done"}.get(priority_of(t, today), ""),
        month_txt=(f"{month['month_current']} מתוך {s['months_total'] or month['month_total']}" if month else "לא נקבע"))


@app.route("/sites/<int:site_id>/meetings/new", methods=["GET", "POST"])
def meeting_upload(site_id):
    if request.method == "POST":
        f = request.files["file"]
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        f.save(tmp.name)
        rows = extract_rows(tmp.name)
        cur = db().execute(
            "INSERT INTO meetings(site_id, kind, source_name, staged_rows) VALUES (?,?,?,?)",
            (site_id, request.form.get("kind", "ביצוע"), f.filename, json.dumps(rows, ensure_ascii=False)))
        db().commit()
        return redirect(url_for("meeting_review", meeting_id=cur.lastrowid))
    return page("העלאת פרוטוקול", """
<form method="post" enctype="multipart/form-data">
<p>סוג פגישה <select name="kind"><option>ביצוע</option><option>שבועית</option><option>סופרפוזיציה</option></select></p>
<input type="file" name="file" accept=".pdf" required> <button>העלאה לבדיקה</button></form>""", url_for=url_for)


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
    return page("בדיקת פרוטוקול לפני פרסום", """
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
""", rows=rows)


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
