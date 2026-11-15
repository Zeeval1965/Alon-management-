"""Smoke and acceptance tests for the Hilel HaZaken 14 app.
Run from the app folder:  python -m unittest discover -s tests -v
The protocol PDFs are read from SAMPLE_DIR (set env HILEL_SAMPLES) and skipped when missing."""
import io
import json
import os
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

import app as A  # noqa: E402
import protocol_parser as P  # noqa: E402
from werkzeug.security import generate_password_hash  # noqa: E402

SAMPLE_DIR = os.environ.get("HILEL_SAMPLES", "/home/claude/in")


class AppTestCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        A.DB_PATH = os.path.join(self.tmp.name, "test.db")
        A.UPLOAD_DIR = os.path.join(self.tmp.name, "uploads")
        A.init_db()
        self.client = A.app.test_client()
        A.app.config["TESTING"] = True
        with A.app.app_context():
            d = A.db()
            d.execute("INSERT INTO sites(name) VALUES ('הלל הזקן 14')")
            people = [("admin@test", 1, None), ("sup@test", 0, "מפקח"),
                      ("owner@test", 0, "בעלי דירות"), ("cons@test", 0, "יועצים")]
            for email, admin, role in people:
                cur = d.execute("INSERT INTO contacts(email, first_name, last_name, password_hash, is_admin) "
                                "VALUES (?,?,?,?,?)", (email, email, "", generate_password_hash("pw"), admin))
                if role:
                    d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,?)",
                              (cur.lastrowid, role))
            d.commit()

    def tearDown(self):
        self.tmp.cleanup()

    def login(self, email):
        c = A.app.test_client()
        c.post("/login", data={"email": email, "password": "pw"})
        return c

    def db_one(self, sql, args=()):
        with A.app.app_context():
            return A.db().execute(sql, args).fetchone()

    def db_val(self, sql, args=()):
        row = self.db_one(sql, args)
        return list(row)[0] if row else None


class AccessTests(AppTestCase):
    def test_anonymous_is_sent_to_login(self):
        r = self.client.get("/sites/1/decisions")
        self.assertEqual(r.status_code, 302)
        self.assertIn("/login", r.headers["Location"])

    def test_role_matrix(self):
        cases = [
            ("admin@test", "/sites/1/supervision", "GET", 200),
            ("sup@test", "/sites/1/supervision", "GET", 200),
            ("owner@test", "/sites/1/supervision", "GET", 403),
            ("owner@test", "/sites/1/decisions", "GET", 200),
            ("owner@test", "/sites/1/plans", "POST", 403),
            ("owner@test", "/sites/1/progress", "GET", 200),
            ("owner@test", "/sites/1/progress", "POST", 403),
            ("cons@test", "/sites/1/plans", "GET", 200),
            ("cons@test", "/sites/1/supervision", "GET", 403),
            ("owner@test", "/admin/queue", "GET", 403),
            ("sup@test", "/admin/queue", "GET", 403),
            ("admin@test", "/admin/queue", "GET", 200),
        ]
        for email, path, method, want in cases:
            c = self.login(email)
            r = c.get(path) if method == "GET" else c.post(path, data={})
            self.assertEqual(r.status_code, want, f"{email} {method} {path}")

    def test_outsider_has_no_site(self):
        with A.app.app_context():
            A.db().execute("INSERT INTO contacts(email, first_name, last_name, password_hash, is_admin) "
                           "VALUES ('out@test','x','',?,0)", (generate_password_hash("pw"),))
            A.db().commit()
        self.assertEqual(self.login("out@test").get("/sites/1").status_code, 403)


class DevRoutingTests(AppTestCase):
    def test_every_message_goes_to_owner_email(self):
        with A.app.app_context():
            d = A.db()
            d.execute("INSERT INTO contacts(email, first_name, last_name, role, pref_channel, mobile) "
                      "VALUES ('g@x.com','גיא','פז','מנהל','whatsapp','050-0000000')")
            cid = d.execute("SELECT id FROM contacts WHERE email='g@x.com'").fetchone()[0]
            d.execute("INSERT INTO memberships(site_id, contact_id, site_role, notify_freq) "
                      "VALUES (1,?,'מנהל אתר','מיידי')", (cid,))
            d.commit()
            A.notify_task_event({"site_id": 1, "public_no": "1", "topic": "x", "text": "בדיקה"}, "בדיקה")
            row = d.execute("SELECT o.channel, o.body, c.email FROM outbox o JOIN contacts c ON c.id=o.contact_id").fetchone()
        self.assertEqual(row["email"], A.DEV_EMAIL)
        self.assertEqual(row["channel"], "email")
        self.assertIn("במקור אל:", row["body"])


@unittest.skipUnless(os.path.exists(os.path.join(SAMPLE_DIR, "f1.pdf")), "f1.pdf not available")
class ProtocolTests(AppTestCase):
    def publish(self, path, kind):
        rows = P.extract_rows(path)
        with A.app.app_context():
            A.db().execute("INSERT INTO meetings(site_id, kind, staged_rows, status) VALUES (1,?,?, 'draft')",
                           (kind, json.dumps(rows, ensure_ascii=False)))
            A.db().commit()
            mid = A.db().execute("SELECT MAX(id) FROM meetings").fetchone()[0]
            A.publish_meeting(mid)
        return rows

    def test_execution_protocol_f1(self):
        rows = self.publish(os.path.join(SAMPLE_DIR, "f1.pdf"), "ביצוע")
        self.assertEqual(len(rows), 42)
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM tasks"), 24)
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM admin_queue"), 0)
        self.assertGreater(self.db_val("SELECT COUNT(*) FROM decisions WHERE source='פרוטוקול'"), 0)

    def test_removal_only_within_same_kind(self):
        self.publish(os.path.join(SAMPLE_DIR, "f1.pdf"), "ביצוע")
        before = self.db_val("SELECT COUNT(*) FROM tasks WHERE status='בוצע'")
        rows = P.extract_rows(os.path.join(SAMPLE_DIR, "f1.pdf"))
        partial = [r for r in rows if r["meeting_date"] == "20.10.26"][:3]
        with A.app.app_context():
            A.db().execute("INSERT INTO meetings(site_id, kind, staged_rows, status) VALUES (1,'ביצוע',?, 'draft')",
                           (json.dumps(partial, ensure_ascii=False),))
            A.db().commit()
            A.publish_meeting(A.db().execute("SELECT MAX(id) FROM meetings").fetchone()[0])
        self.assertGreater(self.db_val("SELECT COUNT(*) FROM tasks WHERE status='בוצע'"), before)

    def test_queue_new_task_action(self):
        with A.app.app_context():
            A.db().execute("INSERT INTO meetings(site_id, kind, staged_rows, status) VALUES (1,'ביצוע','[]','draft')")
            item = {"number": "9", "topic": "חשמל", "task": "בדיקה לתור", "assignee": "", "update": "",
                    "target": "שוטף", "meeting_date": "01.11.26"}
            A.to_queue(1, 1, item, "בדיקה")
            A.db().commit()
            qid = A.db().execute("SELECT id FROM admin_queue").fetchone()[0]
        c = self.login("admin@test")
        r = c.post(f"/admin/queue/{qid}/resolve", data={"action": "new"})
        self.assertEqual(r.status_code, 302)
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM tasks WHERE text='בדיקה לתור'"), 1)


class ModuleTests(AppTestCase):
    def test_decisions_add_and_list(self):
        c = self.login("sup@test")
        c.post("/sites/1/decisions", data={"topic": "חשמל", "text": "אישור גופי תאורה", "decided_on": "20.10.26"})
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM decisions WHERE topic='חשמל'"), 1)

    def test_plan_upload_and_download(self):
        c = self.login("cons@test")
        c.post("/sites/1/plans", data={"title": "חתך א-א", "version": "A",
                                       "file": (io.BytesIO(b"%PDF-1.4 test"), "section-aa.pdf")},
               content_type="multipart/form-data")
        r = c.get("/plans/1/file")
        self.assertEqual(r.status_code, 200)
        self.assertTrue(r.data.startswith(b"%PDF"))

    def test_progress_grid_saves_cell(self):
        c = self.login("sup@test")
        c.post("/sites/1/progress", data={"action": "config", "levels": "קרקע\nקומה 1", "stages": ""})
        c.post("/sites/1/progress", data={"level": "קרקע", "stage": "שלד", "status": "בביצוע"})
        self.assertEqual(self.db_val("SELECT status FROM stage_status WHERE level='קרקע' AND stage='שלד'"), "בביצוע")
        page = self.login("owner@test").get("/sites/1/progress").get_data(as_text=True)
        self.assertIn("קומה 1", page)

    @unittest.skipUnless(os.path.exists(os.path.join(SAMPLE_DIR, "f4.pdf")), "f4.pdf not available")
    def test_supervision_report_f4(self):
        c = self.login("sup@test")
        with open(os.path.join(SAMPLE_DIR, "f4.pdf"), "rb") as fh:
            c.post("/sites/1/supervision", data={"file": (fh, "f4.pdf"), "report_date": "09.11.26"},
                   content_type="multipart/form-data")
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM supervision_items"), 28)


class ContactImportTests(AppTestCase):
    """Contact import reads the first sheet that has a 'דוא"ל' column and tolerates the template layout."""

    def write_xlsx(self, sheets):
        import openpyxl
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for title, rows in sheets:
            ws = wb.create_sheet(title)
            for row in rows:
                ws.append(row)
        path = os.path.join(self.tmp.name, "contacts.xlsx")
        wb.save(path)
        return path

    def contact(self, email):
        with A.app.app_context():
            row = A.db().execute("SELECT first_name, last_name, role FROM contacts WHERE email=?", (email,)).fetchone()
        return tuple(row) if row else None

    def test_reads_second_sheet_when_first_sheet_is_instructions(self):
        path = self.write_xlsx([
            ("הסבר", [["תבנית ייבוא - הלל הזקן 14"], ["הוראות כלליות"]]),
            ("יועצים", [["תחום", "שם מלא", "תפקיד", "דוא\"ל"],
                        ["קרקע", "אביה זאב", "מהנדס קרקע", "aviah@example.com"],
                        ["", "", "", ""]]),
        ])
        with A.app.app_context():
            added, skipped = A.import_contacts(path)
        self.assertEqual((added, skipped), (1, 0))
        self.assertEqual(self.contact("aviah@example.com"), ("אביה", "זאב", "מהנדס קרקע"))

    def test_full_name_is_split_into_first_and_last(self):
        path = self.write_xlsx([("בעלי דירות", [["שם מלא", "דוא\"ל", "סוג"],
                                                ["דנה כהן ולוי", "dana@example.com", "בעל דירה"]])])
        with A.app.app_context():
            A.import_contacts(path)
        self.assertEqual(self.contact("dana@example.com"), ("דנה", "כהן ולוי", "בעל דירה"))

    def test_typographic_quote_in_header_is_recognised(self):
        path = self.write_xlsx([("יועצים", [["שם פרטי", "דוא\u05f4ל"], ["רן", "ran@example.com"]])])
        with A.app.app_context():
            added, _ = A.import_contacts(path)
        self.assertEqual(added, 1)
        self.assertIsNotNone(self.contact("ran@example.com"))

    def test_rows_without_email_are_counted_as_skipped(self):
        path = self.write_xlsx([("יועצים", [["שם פרטי", "דוא\"ל"], ["א", ""], ["ב", "b@example.com"]])])
        with A.app.app_context():
            added, skipped = A.import_contacts(path)
        self.assertEqual((added, skipped), (1, 1))

    def test_missing_email_column_raises_clear_error(self):
        path = self.write_xlsx([("יועצים", [["שם פרטי", "תפקיד"], ["א", "מהנדס"]])])
        with A.app.app_context():
            with self.assertRaises(ValueError):
                A.import_contacts(path)


class ContactPreviewFlowTests(AppTestCase):
    """Upload shows a preview with problems; nothing is imported until the admin confirms."""

    def xlsx_bytes(self, rows):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "יועצים"
        for row in rows:
            ws.append(row)
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def test_preview_then_confirm_imports_only_valid_rows(self):
        data = self.xlsx_bytes([["שם פרטי", "דוא\"ל", "תפקיד"],
                                ["ענת", "anat3@example.com", "יזם"],
                                ["אבי", "not-an-email", "יועצים"]])
        c = self.login("admin@test")
        r = c.post("/contacts/import/preview", data={"site_id": "1",
                                                     "file": (io.BytesIO(data), "c.xlsx")},
                   content_type="multipart/form-data")
        html = r.get_data(as_text=True)
        self.assertEqual(r.status_code, 200)
        self.assertIn("לא תקין", html)
        self.assertIsNone(self.db_one("SELECT id FROM contacts WHERE email='anat3@example.com'"))
        import re as _re
        token = _re.search(r'name="token" value="([0-9a-f]{32})"', html).group(1)
        r2 = c.post("/contacts/import/confirm", data={"token": token, "site_id": "1", "include_0": "on",
                                                      "include_1": "on"})
        self.assertEqual(r2.status_code, 302)
        self.assertIsNotNone(self.db_one("SELECT id FROM contacts WHERE email='anat3@example.com'"))
        self.assertIsNone(self.db_one("SELECT id FROM contacts WHERE email='not-an-email'"))
        self.assertEqual(self.db_one("SELECT m.site_role FROM memberships m JOIN contacts c ON c.id=m.contact_id "
                                     "WHERE c.email='anat3@example.com'")[0], "יזם")
        self.assertFalse(os.path.exists(os.path.join(A.UPLOAD_DIR, "imports", token + ".xlsx")))

    def test_confirm_rejects_malformed_token(self):
        c = self.login("admin@test")
        r = c.post("/contacts/import/confirm", data={"token": "../../etc", "site_id": "1"})
        self.assertEqual(r.status_code, 302)


class ContactMembershipImportTests(AppTestCase):
    """Importing with a site and role links every row to that site as a member."""

    write_xlsx = ContactImportTests.write_xlsx

    def test_import_with_site_and_role_creates_memberships(self):
        path = self.write_xlsx([("יועצים", [["שם פרטי", "דוא\"ל"], ["ענת", "anat@example.com"],
                                            ["אבי", "avi@example.com"]])])
        with A.app.app_context():
            added, _ = A.import_contacts(path, site_id=1, site_role="יועצים")
            rows = A.db().execute(
                "SELECT c.email, m.site_role FROM memberships m JOIN contacts c ON c.id=m.contact_id "
                "WHERE m.site_id=1 AND c.email LIKE '%@example.com' ORDER BY c.email").fetchall()
        self.assertEqual(added, 2)
        self.assertEqual([tuple(r) for r in rows], [("anat@example.com", "יועצים"), ("avi@example.com", "יועצים")])

    def test_row_role_that_is_a_site_role_overrides_form_role(self):
        path = self.write_xlsx([("יועצים", [["שם פרטי", "דוא\"ל", "תפקיד"],
                                            ["יוסי", "yossi@example.com", "יזם"],
                                            ["מיכל", "michal@example.com", "מנהל עבודה"]])])
        with A.app.app_context():
            A.import_contacts(path, site_id=1, site_role="יועצים")
            rows = A.db().execute(
                "SELECT c.email, m.site_role FROM memberships m JOIN contacts c ON c.id=m.contact_id "
                "WHERE m.site_id=1 AND c.email LIKE '%@example.com' ORDER BY c.email").fetchall()
        self.assertEqual([tuple(r) for r in rows],
                         [("michal@example.com", "מנהל עבודה"), ("yossi@example.com", "יזם")])

    def test_role_comes_from_row_with_aliases_and_any_role_text_is_linked(self):
        path = self.write_xlsx([("בעלי דירות", [["שם מלא", "דוא\"ל", "סוג"],
                                                ["דנה כהן", "dana2@example.com", "בעל דירה"],
                                                ["רון לוי", "ron2@example.com", "מנהל עבודה"]])])
        with A.app.app_context():
            added, _ = A.import_contacts(path, site_id=1)
            rows = A.db().execute(
                "SELECT c.email, m.site_role FROM memberships m JOIN contacts c ON c.id=m.contact_id "
                "WHERE m.site_id=1 AND c.email LIKE '%2@example.com' ORDER BY c.email").fetchall()
            contact_count = A.db().execute(
                "SELECT COUNT(*) FROM contacts WHERE email LIKE '%2@example.com'").fetchone()[0]
        self.assertEqual(added, 2)
        self.assertEqual(contact_count, 2)
        self.assertEqual([tuple(r) for r in rows],
                         [("dana2@example.com", "בעלי דירות"), ("ron2@example.com", "מנהל עבודה")])

    def test_reimport_does_not_duplicate_memberships(self):
        path = self.write_xlsx([("יועצים", [["שם פרטי", "דוא\"ל"], ["ענת", "anat@example.com"]])])
        with A.app.app_context():
            A.import_contacts(path, site_id=1, site_role="יועצים")
            A.import_contacts(path, site_id=1, site_role="יועצים")
            n = A.db().execute("SELECT COUNT(*) FROM memberships m JOIN contacts c ON c.id=m.contact_id "
                               "WHERE c.email='anat@example.com' AND m.site_id=1").fetchone()[0]
        self.assertEqual(n, 1)


class TaskAssignmentRoleTests(AppTestCase):
    """One responsible party per task; the others are view-only."""

    def setUp(self):
        super().setUp()
        with A.app.app_context():
            d = A.db()
            d.execute("INSERT INTO tasks(site_id, public_no, text, assignee_text) "
                      "VALUES (1, '1', 'לתקן נזילה', 'ליאור י וגיא')")
            for name in ("ליאור", "גיא"):
                cur = d.execute("INSERT INTO contacts(email, first_name, last_name) VALUES (?,?,?)",
                                (name + "@x.test", name, "י" if name == "ליאור" else ""))
                d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'יועצים')",
                          (cur.lastrowid,))
                d.execute("INSERT INTO task_assignments(task_id, contact_id, raw_name) VALUES (1,?,?)",
                          (cur.lastrowid, name))
            d.commit()
        self.ids = [r[0] for r in self.rows(
            "SELECT contact_id FROM task_assignments WHERE task_id=1 ORDER BY contact_id")]

    def rows(self, sql, args=()):
        with A.app.app_context():
            return [tuple(r) for r in A.db().execute(sql, args).fetchall()]

    def responsible(self):
        return [r[0] for r in self.rows(
            "SELECT contact_id FROM task_assignments WHERE task_id=1 AND is_responsible=1")]

    def test_first_confirmed_person_becomes_responsible(self):
        sup = self.login("sup@test")
        a, b = self.ids
        sup.post(f"/tasks/1/assignees/{a}/confirm")
        self.assertEqual(self.responsible(), [a])
        sup.post(f"/tasks/1/assignees/{b}/confirm")
        self.assertEqual(self.responsible(), [a])  # second person is view-only

    def test_switching_responsible_keeps_exactly_one(self):
        sup = self.login("sup@test")
        a, b = self.ids
        sup.post(f"/tasks/1/assignees/{a}/confirm")
        sup.post(f"/tasks/1/assignees/{b}/confirm")
        sup.post(f"/tasks/1/assignees/{b}/responsible")
        self.assertEqual(self.responsible(), [b])
        self.assertEqual(len(self.responsible()), 1)

    def test_unconfirmed_person_cannot_be_responsible(self):
        sup = self.login("sup@test")
        a, _ = self.ids
        r = sup.post(f"/tasks/1/assignees/{a}/responsible")
        self.assertEqual(r.status_code, 400)
        self.assertEqual(self.responsible(), [])

    def test_only_supervisor_can_change_responsibility(self):
        owner = self.login("owner@test")
        a, _ = self.ids
        r = owner.post(f"/tasks/1/assignees/{a}/responsible")
        self.assertEqual(r.status_code, 403)

    def test_assignments_page_shows_role_labels(self):
        sup = self.login("sup@test")
        a, b = self.ids
        sup.post(f"/tasks/1/assignees/{a}/confirm")
        sup.post(f"/tasks/1/assignees/{b}/confirm")
        html = sup.get("/sites/1/assignments").get_data(as_text=True)
        self.assertIn("אחראי לביצוע", html)
        self.assertIn("לצפייה", html)


class SiteContactsTests(AppTestCase):
    def test_members_see_read_only_contacts_list(self):
        owner = self.login("owner@test")
        r = owner.get("/sites/1/contacts")
        self.assertEqual(r.status_code, 200)
        html = r.get_data(as_text=True)
        self.assertIn("owner@test", html)
        self.assertNotIn("password_hash", html)
        self.assertNotIn("<form", html)  # no editing controls

    def test_main_site_page_links_to_contacts(self):
        owner = self.login("owner@test")
        html = owner.get("/sites/1").get_data(as_text=True)
        self.assertIn("/sites/1/contacts", html)

    def test_contacts_list_is_not_writable(self):
        owner = self.login("owner@test")
        self.assertIn(owner.post("/sites/1/contacts").status_code, (403, 405))


class ReminderRecipientTests(TaskAssignmentRoleTests):
    def test_reminder_goes_to_responsible_only(self):
        sup = self.login("sup@test")
        a, b = self.ids
        sup.post(f"/tasks/1/assignees/{a}/confirm")
        sup.post(f"/tasks/1/assignees/{b}/confirm")
        with A.app.app_context():
            A.db().execute("UPDATE tasks SET target_date=? WHERE id=1", (A.dt.date.today().isoformat(),))
            A.db().commit()
            A.send_task_reminders(1, A.dt.date.today())
        bodies = " ".join(r[0] for r in self.rows("SELECT body FROM outbox WHERE site_id=1"))
        self.assertIn("ליאור", bodies)      # responsible party (ids[0]) gets the reminder
        self.assertNotIn("גיא", bodies)     # view-only person gets no copy

    def test_cc_is_marked_and_reported_but_not_sent(self):
        sup = self.login("sup@test")
        a, b = self.ids
        sup.post(f"/tasks/1/assignees/{a}/confirm")
        sup.post(f"/tasks/1/assignees/{b}/confirm")
        sup.post(f"/tasks/1/assignees/{b}/cc")
        self.assertEqual(self.rows("SELECT is_cc FROM task_assignments WHERE task_id=1 AND contact_id=?", (b,)),
                         [(1,)])
        html = sup.get("/sites/1/cc-report").get_data(as_text=True)
        self.assertIn("גיא", html)
        self.assertIn("<td>1</td>", html)  # counted once
        with A.app.app_context():
            A.db().execute("UPDATE tasks SET target_date=? WHERE id=1", (A.dt.date.today().isoformat(),))
            A.db().commit()
            A.send_task_reminders(1, A.dt.date.today())
        bodies = " ".join(r[0] for r in self.rows("SELECT body FROM outbox WHERE site_id=1"))
        self.assertNotIn("גיא", bodies)

    def test_responsible_cannot_be_marked_cc(self):
        sup = self.login("sup@test")
        a, _ = self.ids
        sup.post(f"/tasks/1/assignees/{a}/confirm")
        self.assertEqual(sup.post(f"/tasks/1/assignees/{a}/cc").status_code, 400)

class ContractorResponsibleTests(TaskAssignmentRoleTests):
    def make_contractor(self):
        with A.app.app_context():
            A.db().execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'קבלן')", (self.ids[0],))
            A.db().commit()

    def test_contractor_task_allows_several_responsible_and_all_are_reminded(self):
        sup = self.login("sup@test")
        a, b = self.ids
        self.make_contractor()
        sup.post(f"/tasks/1/assignees/{a}/confirm")
        sup.post(f"/tasks/1/assignees/{b}/confirm")
        sup.post(f"/tasks/1/assignees/{b}/responsible/add")
        self.assertEqual(sorted(r[0] for r in self.rows(
            "SELECT contact_id FROM task_assignments WHERE task_id=1 AND is_responsible=1")), [a, b])
        with A.app.app_context():
            A.db().execute("UPDATE tasks SET target_date=? WHERE id=1", (A.dt.date.today().isoformat(),))
            A.db().commit()
            A.send_task_reminders(1, A.dt.date.today())
        bodies = " ".join(r[0] for r in self.rows("SELECT body FROM outbox WHERE site_id=1"))
        self.assertIn("ליאור", bodies)
        self.assertIn("גיא", bodies)

    def test_second_responsible_refused_without_contractor(self):
        sup = self.login("sup@test")
        a, b = self.ids
        sup.post(f"/tasks/1/assignees/{a}/confirm")
        sup.post(f"/tasks/1/assignees/{b}/confirm")
        self.assertEqual(sup.post(f"/tasks/1/assignees/{b}/responsible/add").status_code, 400)
        self.assertEqual(len(self.rows(
            "SELECT 1 FROM task_assignments WHERE task_id=1 AND is_responsible=1")), 1)


class ContactManagementTests(AppTestCase):
    def setUp(self):
        super().setUp()
        with A.app.app_context():  # real-looking addresses: the email check requires a domain dot
            for old, new in (("sup@test", "sup@x.test"), ("owner@test", "owner@x.test"), ("admin@test", "admin@x.test")):
                A.db().execute("UPDATE contacts SET email=? WHERE email=?", (new, old))
            A.db().commit()

    def owner_id(self):
        return self.db_val("SELECT id FROM contacts WHERE email='owner@x.test'")

    def test_supervisor_can_edit_but_not_blank_a_field(self):
        sup = self.login("sup@x.test")
        cid = self.owner_id()
        sup.post(f"/contacts/{cid}/edit", data={"phone": "03-1111111", "email": "owner@x.test"})
        self.assertEqual(self.db_val("SELECT phone FROM contacts WHERE id=?", (cid,)), "03-1111111")
        sup.post(f"/contacts/{cid}/edit", data={"phone": "", "email": "owner@x.test"})
        self.assertEqual(self.db_val("SELECT phone FROM contacts WHERE id=?", (cid,)), "03-1111111")

    def test_administrator_can_blank_a_field(self):
        admin = self.login("admin@x.test")
        cid = self.owner_id()
        admin.post(f"/contacts/{cid}/edit", data={"phone": "", "email": "owner@x.test"})
        self.assertIsNone(self.db_val("SELECT phone FROM contacts WHERE id=?", (cid,)))

    def test_owner_cannot_edit_or_import(self):
        owner = self.login("owner@x.test")
        cid = self.owner_id()
        self.assertEqual(owner.get(f"/contacts/{cid}/edit").status_code, 403)
        self.assertEqual(owner.get("/contacts/import").status_code, 403)

    def test_supervisor_can_open_import(self):
        self.assertEqual(self.login("sup@x.test").get("/contacts/import").status_code, 200)

    def test_duplicate_email_is_rejected(self):
        sup = self.login("sup@x.test")
        cid = self.owner_id()
        r = sup.post(f"/contacts/{cid}/edit", data={"email": "sup@x.test"})
        self.assertIn("כבר שייכת", r.get_data(as_text=True))
        self.assertEqual(self.db_val("SELECT email FROM contacts WHERE id=?", (cid,)), "owner@x.test")

    def test_list_shows_edit_only_to_managers(self):
        html_sup = self.login("sup@x.test").get("/sites/1/contacts").get_data(as_text=True)
        html_owner = self.login("owner@x.test").get("/sites/1/contacts").get_data(as_text=True)
        self.assertIn("/contacts/", html_sup)
        self.assertNotIn("/edit", html_owner)


class ContactRoleAndDeleteTests(ContactManagementTests):
    def test_supervisor_can_add_role_but_not_remove_it(self):
        sup = self.login("sup@x.test")
        cid = self.owner_id()
        sup.post(f"/contacts/{cid}/roles/add", data={"site_id": "1", "site_role": "קבלן ראשי"})
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM memberships WHERE contact_id=? AND site_role='קבלן'", (cid,)), 1)
        r = sup.post(f"/contacts/{cid}/roles/remove", data={"site_id": "1", "site_role": "קבלן"})
        self.assertEqual(r.status_code, 403)

    def test_administrator_can_remove_role(self):
        admin = self.login("admin@x.test")
        cid = self.owner_id()
        admin.post(f"/contacts/{cid}/roles/remove", data={"site_id": "1", "site_role": "בעלי דירות"})
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM memberships WHERE contact_id=?", (cid,)), 0)

    def test_only_administrator_can_delete_and_not_self_or_admin(self):
        cid = self.owner_id()
        self.assertEqual(self.login("sup@x.test").post(f"/contacts/{cid}/delete").status_code, 403)
        admin = self.login("admin@x.test")
        self.assertEqual(admin.post("/contacts/1/delete").status_code, 400)  # own account
        admin.post(f"/contacts/{cid}/delete")
        self.assertIsNone(self.db_val("SELECT id FROM contacts WHERE id=?", (cid,)))
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM memberships WHERE contact_id=?", (cid,)), 0)


class AutoAssignTests(AppTestCase):
    def setUp(self):
        super().setUp()
        with A.app.app_context():
            d = A.db()
            cur = d.execute("INSERT INTO contacts(email, first_name, last_name) VALUES ('gil@x.test','גיל','כהן')")
            d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'קבלן')", (cur.lastrowid,))
            dan = d.execute("INSERT INTO contacts(email, first_name, last_name) VALUES ('dan@x.test','דני','לוי')")
            d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'יועצים')", (dan.lastrowid,))
            d.commit()

    def add_task(self, text):
        with A.app.app_context():
            A.db().execute("INSERT INTO tasks(site_id, public_no, text, assignee_text) VALUES (1, ?, 'x', ?)",
                           (str(len(text)) + text[:3], text))
            tid = A.db().execute("SELECT MAX(id) FROM tasks").fetchone()[0]
            A.suggest_assignees(tid)
            A.db().commit()
            return tid

    def test_full_name_is_linked_and_confirmed_as_responsible(self):
        tid = self.add_task("גיל כהן")
        row = self.db_one("SELECT confirmed, is_responsible FROM task_assignments WHERE task_id=?", (tid,))
        self.assertEqual((row["confirmed"], row["is_responsible"]), (1, 1))

    def test_partial_name_waits_for_supervisor(self):
        tid = self.add_task("דני")
        row = self.db_one("SELECT confirmed, is_responsible FROM task_assignments WHERE task_id=?", (tid,))
        self.assertEqual((row["confirmed"], row["is_responsible"]), (0, 0))


if __name__ == "__main__":
    unittest.main()
