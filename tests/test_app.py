"""Smoke and acceptance tests for the Hilel HaZaken 14 app.
Run from the app folder:  python -m unittest discover -s tests -v
The protocol PDFs are read from SAMPLE_DIR (set env HILEL_SAMPLES) and skipped when missing."""
import io
import re
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
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM tasks WHERE status='בוצע'"), before)
        self.assertGreater(self.db_val("SELECT COUNT(*) FROM task_closure_pending"), 0)

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

    def test_partial_name_is_linked_when_unique_on_the_site(self):
        tid = self.add_task("דני")
        row = self.db_one("SELECT confirmed, is_responsible FROM task_assignments WHERE task_id=?", (tid,))
        self.assertEqual((row["confirmed"], row["is_responsible"]), (1, 1))

    def test_partial_name_waits_when_several_people_share_the_first_name(self):
        with A.app.app_context():
            d = A.db()
            cur = d.execute("INSERT INTO contacts(email, first_name, last_name) VALUES ('dani2@x.test','דני','כהן')")
            d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'יועצים')", (cur.lastrowid,))
            d.commit()
        tid = self.add_task("דני")
        rows = [tuple(r) for r in self.rows_for(tid)]
        self.assertTrue(rows)
        self.assertTrue(all(r[0] == 0 for r in rows))

    def rows_for(self, tid):
        with A.app.app_context():
            return A.db().execute("SELECT confirmed FROM task_assignments WHERE task_id=?", (tid,)).fetchall()


class MultiRoleCellTests(AppTestCase):
    def test_role_cell_with_several_roles_becomes_several_memberships(self):
        with A.app.app_context():
            A.import_rows([{"email": "multi@x.test", "first": "אביה", "last": "זאבי",
                            "role": "יועץ קרקע, מפקח"}], site_id=1)
            rows = A.db().execute("SELECT m.site_role FROM memberships m JOIN contacts c ON c.id=m.contact_id "
                                  "WHERE c.email='multi@x.test' ORDER BY m.site_role").fetchall()
            roles = sorted(r[0] for r in rows)
        self.assertEqual(roles, ["יועץ קרקע", "מפקח"])


class WarningsTests(AppTestCase):
    def test_same_person_same_role_twice_in_file_is_flagged_for_deletion(self):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "אנשי קשר"
        ws.append(['דוא"ל', "שם פרטי", "שם משפחה", "תפקיד", "ארגון"])
        ws.append(["a@x.test", "דני", "לוי", "נציגות", "א"])
        ws.append(["a@x.test", "דני", "לוי", "נציגות", "א"])
        path = os.path.join(self.tmp.name, "dup.xlsx")
        wb.save(path)
        rows = A._contact_rows(path)
        self.assertIn("כפילות זהה", rows[1]["notes"][0])

    def test_warnings_page_lists_duplicates_and_unassigned_tasks(self):
        with A.app.app_context():
            d = A.db()
            d.execute("INSERT INTO contacts(email, first_name, last_name) VALUES ('d1@x.test','דני','לוי')")
            d.execute("INSERT INTO contacts(email, first_name, last_name) VALUES ('d2@x.test','דני','לוי')")
            for cid in d.execute("SELECT id FROM contacts WHERE email IN ('d1@x.test','d2@x.test')").fetchall():
                d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'יועצים')", (cid[0],))
            d.execute("INSERT INTO tasks(site_id, public_no, text, assignee_text) VALUES (1,'9','משימה ללא שיוך','ישראל')")
            d.commit()
        html = self.login("sup@test").get("/sites/1/warnings").get_data(as_text=True)
        self.assertIn("d1@x.test", html)
        self.assertIn("משימה ללא שיוך", html)


class RolesSyncTests(AppTestCase):
    def make_file(self, rows):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "אנשי קשר"
        ws.append(['דוא"ל', "שם פרטי", "שם משפחה", "תפקיד", "ארגון"])
        for r in rows:
            ws.append(r)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def test_sync_sets_roles_to_file_and_leaves_others_alone(self):
        with A.app.app_context():
            d = A.db()
            cur = d.execute("INSERT INTO contacts(email, first_name, last_name) VALUES ('yoav@x.test','יואב','ש')")
            d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'מפקח')", (cur.lastrowid,))
            d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'יועץ קרקע')", (cur.lastrowid,))
            d.commit()
        admin = self.login("admin@test")
        buf = self.make_file([["yoav@x.test", "יואב", "ש", "יועץ קרקע", "א"],
                              ["nobody@x.test", "לא", "קיים", "יועץ", "ב"]])
        r = admin.post("/roles/sync/preview", data={"site_id": "1", "file": (buf, "c.xlsx")},
                       content_type="multipart/form-data")
        html = r.get_data(as_text=True)
        self.assertIn("ייעודכן", html)
        self.assertIn("לא קיים במערכת", html)
        token = re.search(r'name="token" value="([0-9a-f]{32})"', html).group(1)
        admin.post("/roles/sync/confirm", data={"token": token, "site_id": "1"})
        with A.app.app_context():
            roles = [r[0] for r in A.db().execute(
                "SELECT m.site_role FROM memberships m JOIN contacts c ON c.id=m.contact_id "
                "WHERE c.email='yoav@x.test' AND m.site_id=1").fetchall()]
        self.assertEqual(roles, ["יועץ קרקע"])

    def test_only_administrator_can_sync(self):
        self.assertEqual(self.login("sup@test").get("/roles/sync").status_code, 403)


class RematchTests(AutoAssignTests):
    def test_name_followed_by_punctuation_matches(self):
        tid = self.add_task("גיל כהן' ודני")
        row = self.db_one("SELECT confirmed, is_responsible FROM task_assignments WHERE task_id=?", (tid,))
        self.assertEqual((row["confirmed"], row["is_responsible"]), (1, 1))

    def test_task_created_before_contact_is_matched_by_rematch(self):
        with A.app.app_context():
            A.db().execute("INSERT INTO tasks(site_id, public_no, text, assignee_text) VALUES (1,'7','x','משה לוי')")
            A.db().commit()
            tid = A.db().execute("SELECT id FROM tasks WHERE public_no='7'").fetchone()[0]
            d = A.db()
            cur = d.execute("INSERT INTO contacts(email, first_name, last_name) VALUES ('moshe@x.test','משה','לוי')")
            d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'יועצים')", (cur.lastrowid,))
            d.commit()
        self.login("sup@test").post("/sites/1/assignments/rematch")
        row = self.db_one("SELECT confirmed FROM task_assignments WHERE task_id=?", (tid,))
        self.assertEqual(row["confirmed"], 1)


class ContactCardTests(AutoAssignTests):
    def test_task_shows_linked_name_as_link_to_card(self):
        tid = self.add_task("גיל כהן")
        html = self.login("sup@test").get("/sites/1").get_data(as_text=True)
        cid = self.db_val("SELECT contact_id FROM task_assignments WHERE task_id=?", (tid,))
        self.assertIn(f"/sites/1/contacts/{cid}", html)
        self.assertIn("גיל כהן", html)

    def test_card_is_read_only_and_visible_to_members(self):
        tid = self.add_task("גיל כהן")
        cid = self.db_val("SELECT contact_id FROM task_assignments WHERE task_id=?", (tid,))
        r = self.login("owner@test").get(f"/sites/1/contacts/{cid}")
        self.assertEqual(r.status_code, 200)
        self.assertIn("כרטיס איש קשר", r.get_data(as_text=True))
        self.assertNotIn("/edit", r.get_data(as_text=True))


class BulkDeleteTests(AppTestCase):
    def test_admin_deletes_selected_and_keeps_self_and_admins(self):
        owner = self.db_val("SELECT id FROM contacts WHERE email='owner@test'")
        admin = self.db_val("SELECT id FROM contacts WHERE email='admin@test'")
        cons = self.db_val("SELECT id FROM contacts WHERE email='cons@test'")
        c = self.login("admin@test")
        c.post("/contacts/bulk-delete", data={"site_id": "1", "ids": [str(owner), str(admin), str(cons)]})
        self.assertIsNone(self.db_val("SELECT id FROM contacts WHERE id=?", (owner,)))
        self.assertIsNone(self.db_val("SELECT id FROM contacts WHERE id=?", (cons,)))
        self.assertIsNotNone(self.db_val("SELECT id FROM contacts WHERE id=?", (admin,)))

    def test_others_cannot_bulk_delete(self):
        cons = self.db_val("SELECT id FROM contacts WHERE email='cons@test'")
        r = self.login("sup@test").post("/contacts/bulk-delete", data={"site_id": "1", "ids": [str(cons)]})
        self.assertEqual(r.status_code, 403)
        self.assertIsNotNone(self.db_val("SELECT id FROM contacts WHERE id=?", (cons,)))

    def test_admin_sees_selection_controls_others_do_not(self):
        html_admin = self.login("admin@test").get("/sites/1/contacts").get_data(as_text=True)
        html_sup = self.login("sup@test").get("/sites/1/contacts").get_data(as_text=True)
        self.assertIn("מחיקת הנבחרים", html_admin)
        self.assertNotIn("מחיקת הנבחרים", html_sup)


class AmbiguousFirstNameTests(AutoAssignTests):
    def test_initial_match_does_not_also_suggest_bare_first_name(self):
        with A.app.app_context():
            d = A.db()
            cur = d.execute("INSERT INTO contacts(email, first_name, last_name) VALUES ('lior@x.test','דני','ישראלי')")
            d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'יועצים')", (cur.lastrowid,))
            d.commit()
        tid = self.add_task("דני י...")
        rows = self.rows_for_names(tid)
        self.assertEqual(rows, ["דני ישראלי"])

    def rows_for_names(self, tid):
        with A.app.app_context():
            return [r[0] for r in A.db().execute(
                "SELECT c.first_name || ' ' || COALESCE(c.last_name,'') FROM task_assignments a "
                "JOIN contacts c ON c.id=a.contact_id WHERE a.task_id=?", (tid,)).fetchall()]


class InformationOnlyTests(AppTestCase):
    def setUp(self):
        super().setUp()
        with A.app.app_context():
            d = A.db()
            d.execute("INSERT INTO tasks(site_id, public_no, text, assignee_text, status, target_date) "
                      "VALUES (1, '5', 'לידיעה בלבד', 'ישראל', 'לידיעה', ?)", (A.dt.date.today().isoformat(),))
            d.commit()

    def test_information_only_task_gets_no_reminder_and_no_warning(self):
        with A.app.app_context():
            A.send_task_reminders(1, A.dt.date.today())
            self.assertEqual(A.db().execute("SELECT COUNT(*) FROM outbox").fetchone()[0], 0)
        html = self.login("sup@test").get("/sites/1/warnings").get_data(as_text=True)
        self.assertNotIn("לידיעה בלבד", html)

    def test_information_only_task_sits_under_information_heading_not_open(self):
        html = self.login("sup@test").get("/sites/1").get_data(as_text=True)
        info_at = html.index("<h3>לידיעה</h3>")
        self.assertGreater(html.index("לידיעה בלבד"), info_at)
        self.assertLess(html.index("<h3>משימות פתוחות</h3>"), info_at)

    def test_status_can_be_set_to_information_only(self):
        with A.app.app_context():
            tid = A.db().execute("SELECT id FROM tasks WHERE public_no='5'").fetchone()[0]
        self.login("sup@test").post(f"/tasks/{tid}/status", data={"status": "בטיפול", "by": "x"})
        self.login("sup@test").post(f"/tasks/{tid}/status", data={"status": "לידיעה", "by": "x"})
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (tid,)), "לידיעה")


class OpenTasksSortTests(AppTestCase):
    def setUp(self):
        super().setUp()
        with A.app.app_context():
            d = A.db()
            d.execute("INSERT INTO tasks(site_id, public_no, topic, text, status) VALUES (1, 'S90', 'ת-נושא', 'א', 'בטיפול')")
            d.execute("INSERT INTO tasks(site_id, public_no, topic, text, status) VALUES (1, 'S91', 'א-נושא', 'ב', 'בטיפול')")
            d.commit()

    def test_sort_by_topic_orders_open_tasks(self):
        html = self.login("sup@test").get("/sites/1?sort=topic").get_data(as_text=True)
        self.assertLess(html.index(">א-נושא<"), html.index(">ת-נושא<"))

    def test_sort_by_date_added_newest_first(self):
        with A.app.app_context():
            d = A.db()
            d.execute("UPDATE tasks SET created_at='2026-01-01' WHERE public_no='S90'")
            d.execute("UPDATE tasks SET created_at='2026-02-01' WHERE public_no='S91'")
            d.commit()
        html = self.login("sup@test").get("/sites/1?sort=added").get_data(as_text=True)
        self.assertLess(html.index(">ב<"), html.index(">א<"))

    def test_sort_control_is_shown(self):
        html = self.login("sup@test").get("/sites/1").get_data(as_text=True)
        self.assertIn('name="sort"', html)


class ProtocolAuditTests(AppTestCase):
    def _meeting(self, kind, day, task, source):
        rows = [{"number": "1", "topic": "נושא", "task": task, "assignee": "", "target": "", "meeting_date": day}]
        with A.app.app_context():
            d = A.db()
            d.execute("INSERT INTO meetings(site_id, kind, source_name, staged_rows, status) VALUES (1,?,?,?,'published')",
                      (kind, source, A.json.dumps(rows, ensure_ascii=False)))
            d.commit()

    def test_weekly_gap_and_duplicates_are_reported(self):
        self._meeting("שבועית", "01.11.26", "משימה א", "w1.pdf")
        self._meeting("שבועית", "08.11.26", "משימה ב", "w2.pdf")
        self._meeting("שבועית", "10.11.26", "משימה ג", "w3.pdf")
        self._meeting("שבועית", "08.11.26", "משימה ב", "w2-copy.pdf")
        html = self.login("sup@test").get("/sites/1/audit").get_data(as_text=True)
        self.assertIn("פער של 2 ימים", html)
        self.assertIn("כפילות תאריך", html)
        self.assertIn("תוכן זהה", html)

    def test_clean_weekly_sequence_is_ok(self):
        self._meeting("שבועית", "01.11.26", "משימה א", "w1.pdf")
        self._meeting("שבועית", "08.11.26", "משימה ב", "w2.pdf")
        html = self.login("sup@test").get("/sites/1/audit").get_data(as_text=True)
        self.assertNotIn("פער של", html)
        self.assertIn("תקין", html)


class ClosureApprovalTests(AppTestCase):
    def setUp(self):
        super().setUp()
        with A.app.app_context():
            d = A.db()
            d.execute("INSERT INTO tasks(site_id, public_no, topic, text, status) VALUES (1,'C1','חשמל','משימה לסגירה','בטיפול')")
            tid = d.execute("SELECT id FROM tasks WHERE public_no='C1'").fetchone()[0]
            d.execute("INSERT INTO task_closure_pending(task_id, meeting_id, created_at) VALUES (?,1,'2026-11-01')", (tid,))
            d.commit()
        with A.app.app_context():
            self.tid = A.db().execute("SELECT id FROM tasks WHERE public_no='C1'").fetchone()[0]

    def test_approve_closes_and_logs_who(self):
        self.login("sup@test").post(f"/tasks/{self.tid}/closure/approve")
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (self.tid,)), "בוצע")
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM task_closure_pending"), 0)
        self.assertEqual(self.db_val("SELECT changed_by FROM task_status_log WHERE task_id=?", (self.tid,)), "sup@test")

    def test_keep_leaves_task_open(self):
        self.login("sup@test").post(f"/tasks/{self.tid}/closure/keep")
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (self.tid,)), "בטיפול")
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM task_closure_pending"), 0)

    def test_owner_cannot_decide_closure(self):
        r = self.login("owner@test").post(f"/tasks/{self.tid}/closure/approve")
        self.assertEqual(r.status_code, 403)
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (self.tid,)), "בטיפול")

    def test_pending_closure_shown_on_warnings_page(self):
        html = self.login("sup@test").get("/sites/1/warnings").get_data(as_text=True)
        self.assertIn("לבדיקה: לא הופיעה בפרוטוקול המלא האחרון", html)
        self.assertIn("סגור כבוצע", html)


class StatusLogTests(AppTestCase):
    def test_status_change_records_user_and_time(self):
        with A.app.app_context():
            A.db().execute("INSERT INTO tasks(site_id, public_no, text, status) VALUES (1,'L1','בדיקה','בטיפול')")
            A.db().commit()
            tid = A.db().execute("SELECT id FROM tasks WHERE public_no='L1'").fetchone()[0]
        self.login("sup@test").post(f"/tasks/{tid}/status", data={"status": "ממתינה לאישור", "by": "גיא"})
        row = self.db_one("SELECT status, handled_by, changed_by, changed_at FROM task_status_log WHERE task_id=?", (tid,))
        self.assertEqual(row["status"], "ממתינה לאישור")
        self.assertEqual(row["handled_by"], "גיא")
        self.assertEqual(row["changed_by"], "sup@test")
        self.assertTrue(row["changed_at"])


class EmailVisibilityTests(AppTestCase):
    def test_owner_does_not_see_participant_emails(self):
        c = self.login("owner@test")
        self.assertNotIn('<th>דוא"ל</th>', c.get("/sites/1/contacts").get_data(as_text=True))
        with A.app.app_context():
            cid = A.db().execute("SELECT id FROM contacts WHERE email='sup@test'").fetchone()[0]
        self.assertNotIn('דוא"ל</th>', c.get(f"/sites/1/contacts/{cid}").get_data(as_text=True))

    def test_supervisor_sees_participant_emails(self):
        html = self.login("sup@test").get("/sites/1/contacts").get_data(as_text=True)
        self.assertIn('<th>דוא"ל</th>', html)
        self.assertIn("<td>owner@test</td>", html)


class BacklogDataTests(AppTestCase):
    def _publish_rows(self, kind, rows):
        with A.app.app_context():
            A.db().execute("INSERT INTO meetings(site_id, kind, staged_rows, status) VALUES (1,?,?, 'draft')",
                           (kind, json.dumps(rows, ensure_ascii=False)))
            A.db().commit()
            mid = A.db().execute("SELECT MAX(id) FROM meetings").fetchone()[0]
            A.publish_meeting(mid)

    def test_global_lists_are_seeded(self):
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM severity_levels"), 4)
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM item_types"), 7)
        self.assertEqual(self.db_val("SELECT kind FROM task_status_options WHERE name='בוטלה'"), "closed")

    def test_suggestions_from_text(self):
        self.assertEqual(A.suggest_severity("תיקון דחוף בחזית"), "גבוה")
        self.assertEqual(A.suggest_severity("סיכון בטיחות"), "קריטי")
        self.assertEqual(A.suggest_item_type("בדיקת חזית הבניין"), "נושא חזית")
        self.assertEqual(A.suggest_item_type("אין התאמה למפרט"), "בקשה למפרט")

    def test_unknown_status_is_rejected(self):
        with A.app.app_context():
            A.db().execute("INSERT INTO tasks(site_id, public_no, text, status) VALUES (1,'S70','בדיקה','פתוחה')")
            A.db().commit()
            tid = A.db().execute("SELECT id FROM tasks WHERE public_no='S70'").fetchone()[0]
        c = self.login("sup@test")
        self.assertEqual(c.post(f"/tasks/{tid}/status", data={"status": "שגיאה"}).status_code, 400)
        self.assertEqual(c.post(f"/tasks/{tid}/status", data={"status": "ממתין למפרט"}).status_code, 302)
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (tid,)), "ממתין למפרט")

    def test_closed_task_reappearing_is_reopened_and_counted(self):
        row = {"number": "1", "topic": "חשמל", "task": "החלפת לוח חשמל קריטי", "assignee": "",
               "target": "", "meeting_date": "01.11.26"}
        self._publish_rows("שבועית", [row])
        tid = self.db_val("SELECT id FROM tasks WHERE public_no LIKE '%#1'")
        self.assertEqual(self.db_val("SELECT severity FROM tasks WHERE id=?", (tid,)), "קריטי")
        with A.app.app_context():
            A.db().execute("UPDATE tasks SET status='בוצע' WHERE id=?", (tid,))
            A.db().commit()
        self._publish_rows("שבועית", [dict(row, meeting_date="08.11.26")])
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (tid,)), "לבדיקה חוזרת")
        self.assertEqual(self.db_val(
            "SELECT COUNT(DISTINCT meeting_id) FROM task_targets WHERE task_id=?", (tid,)), 2)

    def test_kpis_and_recommendations(self):
        row = {"number": "1", "topic": "חשמל", "task": "החלפת לוח חשמל קריטי", "assignee": "",
               "target": "", "meeting_date": "01.11.26"}
        self._publish_rows("שבועית", [row])
        self._publish_rows("שבועית", [dict(row, meeting_date="08.11.26")])
        with A.app.app_context():
            k = A.site_kpis(1)
            recs = A.site_recommendations(1)
        self.assertEqual(k["open"], 1)
        self.assertEqual(k["repeats_open"], 1)
        self.assertEqual(k["critical_open"], 1)
        self.assertTrue(any("לסגור סופית" in r for r in recs))

    def test_comment_and_document_by_supervisor_only(self):
        with A.app.app_context():
            A.db().execute("INSERT INTO tasks(site_id, public_no, text, status) VALUES (1,'S71','בדיקה','פתוחה')")
            A.db().commit()
            tid = A.db().execute("SELECT id FROM tasks WHERE public_no='S71'").fetchone()[0]
        self.login("sup@test").post(f"/tasks/{tid}/comments", data={"text": "הערה לבדיקה"})
        self.login("sup@test").post(f"/tasks/{tid}/docs", data={"name": "תכנית ב-3"})
        self.assertEqual(self.db_val("SELECT text FROM task_comments WHERE task_id=?", (tid,)), "הערה לבדיקה")
        self.assertEqual(self.db_val("SELECT name FROM task_docs WHERE task_id=?", (tid,)), "תכנית ב-3")
        self.assertEqual(self.login("owner@test").post(f"/tasks/{tid}/comments", data={"text": "x"}).status_code, 403)


class TaskLinkTests(AppTestCase):
    def setUp(self):
        super().setUp()
        with A.app.app_context():
            d = A.db()
            owner = d.execute("SELECT id FROM contacts WHERE email='owner@test'").fetchone()[0]
            for no, text in (("K1", "תיקון חזית צפון"), ("K2", "החלפת דלת")):
                d.execute("INSERT INTO tasks(site_id, public_no, topic, text, status) VALUES (1,?,'חזית',?,'פתוחה')", (no, text))
            d.commit()
            self.t1 = d.execute("SELECT id FROM tasks WHERE public_no='K1'").fetchone()[0]
            self.t2 = d.execute("SELECT id FROM tasks WHERE public_no='K2'").fetchone()[0]
            self.owner = owner
            d.execute("INSERT INTO task_assignments(task_id, contact_id, confirmed, is_responsible) VALUES (?,?,1,1)",
                      (self.t1, owner))
            d.commit()
            self.token = A.make_task_link(self.t1, owner)
            A.db().commit()

    def test_link_shows_and_changes_only_its_task(self):
        c = A.app.test_client()
        self.assertIn("תיקון חזית צפון", c.get(f"/u/{self.token}").get_data(as_text=True))
        self.assertNotIn("החלפת דלת", c.get(f"/u/{self.token}").get_data(as_text=True))
        c.post(f"/u/{self.token}", data={"status": "בטיפול", "comment": "הזמנתי חומרים"})
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (self.t1,)), "בטיפול")
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (self.t2,)), "פתוחה")
        self.assertEqual(self.db_val("SELECT changed_by FROM task_status_log WHERE task_id=?", (self.t1,)), "owner@test")
        self.assertEqual(self.db_val("SELECT text FROM task_comments WHERE task_id=?", (self.t1,)), "הזמנתי חומרים")

    def test_unknown_or_expired_link_is_not_found(self):
        c = A.app.test_client()
        self.assertEqual(c.get("/u/not-a-real-token").status_code, 404)
        with A.app.app_context():
            A.db().execute("UPDATE task_links SET expires_at='2000-01-01T00:00' WHERE contact_id=?", (self.owner,))
            A.db().commit()
        self.assertEqual(A.app.test_client().get(f"/u/{self.token}").status_code, 404)

    def test_closing_from_link_is_a_request_only(self):
        A.app.test_client().post(f"/u/{self.token}", data={"status": "בוצע"})
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (self.t1,)), "פתוחה")
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM task_closure_pending WHERE task_id=?", (self.t1,)), 1)

    def test_send_screen_is_for_managers_only(self):
        self.assertEqual(self.login("owner@test").get("/sites/1/send-links").status_code, 403)
        self.assertEqual(self.login("sup@test").get("/sites/1/send-links").status_code, 200)

    def test_send_queues_links_to_owner_in_development(self):
        c = self.login("sup@test")
        r = c.post("/sites/1/send-links", data={f"ch_{self.t1}_{self.owner}": "email"})
        self.assertEqual(r.status_code, 200)
        body = self.db_val("SELECT body FROM outbox ORDER BY id DESC LIMIT 1")
        self.assertIn("/u/", body)
        self.assertIn("[פיתוח", body)
        self.assertGreaterEqual(self.db_val("SELECT COUNT(*) FROM task_links WHERE task_id=?", (self.t1,)), 2)


class DecisionChangesTests(AppTestCase):
    def _task(self, no, status="פתוחה"):
        with A.app.app_context():
            A.db().execute("INSERT INTO tasks(site_id, public_no, topic, text, status) VALUES (1,?,'חשמל',?,?)",
                           (no, "משימה " + no, status))
            A.db().commit()
            return A.db().execute("SELECT id FROM tasks WHERE public_no=?", (no,)).fetchone()[0]

    def test_new_statuses_exist_and_old_name_is_gone(self):
        self.assertEqual(self.db_val("SELECT kind FROM task_status_options WHERE name='מעוכבת'"), "open")
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM task_status_options WHERE name='לא רלוונטי'"), 0)
        tid = self._task("D1")
        c = self.login("sup@test")
        self.assertEqual(c.post(f"/tasks/{tid}/status", data={"status": "ממתינה לבדיקה"}).status_code, 302)
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (tid,)), "ממתינה לבדיקה")

    def test_manager_closing_sets_closure_date(self):
        tid = self._task("D2")
        self.login("sup@test").post(f"/tasks/{tid}/status", data={"status": "בוצע"})
        self.assertEqual(self.db_val("SELECT closed_at FROM tasks WHERE id=?", (tid,)), __import__("datetime").date.today().isoformat())

    def test_confirmed_assignee_reports_but_cannot_close(self):
        tid = self._task("D3")
        with A.app.app_context():
            A.db().execute("INSERT INTO contacts(email, first_name, last_name, password_hash, is_admin) VALUES ('kab@test','קבלן','',?,0)",
                           (A.generate_password_hash("pw"),))
            cid = A.db().execute("SELECT id FROM contacts WHERE email='kab@test'").fetchone()[0]
            A.db().execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'קבלן')", (cid,))
            A.db().execute("INSERT INTO task_assignments(task_id, contact_id, confirmed) VALUES (?,?,1)", (tid, cid))
            A.db().commit()
        c = self.login("kab@test")
        self.assertEqual(c.post(f"/tasks/{tid}/status", data={"status": "בטיפול"}).status_code, 302)
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (tid,)), "בטיפול")
        c.post(f"/tasks/{tid}/status", data={"status": "בוצע"})
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (tid,)), "בטיפול")
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM task_closure_pending WHERE task_id=?", (tid,)), 1)
        other = self._task("D4")
        self.assertEqual(c.post(f"/tasks/{other}/status", data={"status": "בטיפול"}).status_code, 403)

    def test_protocol_search_finds_non_task_rows(self):
        rows = [{"number": "7", "topic": "כללי", "task": "הערה על תיאום מערכות", "assignee": "",
                 "target": "", "meeting_date": "03.11.26", "skip": False}]
        with A.app.app_context():
            A.db().execute("INSERT INTO meetings(site_id, kind, staged_rows, status) VALUES (1,'שבועית',?, 'published')",
                           (json.dumps(rows, ensure_ascii=False),))
            A.db().commit()
        html = self.login("sup@test").get("/sites/1/protocol-search?q=תיאום").get_data(as_text=True)
        self.assertIn("הערה על תיאום מערכות", html)
        self.assertEqual(self.login("owner@test").get("/sites/1/protocol-search?q=תיאום").status_code, 403)

    def test_opened_date_is_protocol_date_and_occurrences_kept(self):
        row = {"number": "1", "topic": "חשמל", "task": "בדיקת לוח חשמל", "assignee": "",
               "target": "", "meeting_date": "01.11.26"}
        with A.app.app_context():
            A.db().execute("INSERT INTO meetings(site_id, kind, staged_rows, status) VALUES (1,'שבועית',?, 'draft')",
                           (json.dumps([row], ensure_ascii=False),))
            A.db().commit()
            A.publish_meeting(A.db().execute("SELECT MAX(id) FROM meetings").fetchone()[0])
        self.assertEqual(self.db_val("SELECT created_at FROM tasks WHERE text='בדיקת לוח חשמל'"), "2026-11-01")
        self.assertEqual(self.db_val("SELECT text FROM task_occurrences WHERE task_id=(SELECT id FROM tasks WHERE text='בדיקת לוח חשמל')"),
                         "בדיקת לוח חשמל")


class SiteDashboardTests(AppTestCase):
    def test_dashboard_shows_kpis_status_bars_and_site_image(self):
        with A.app.app_context():
            A.db().execute("UPDATE sites SET name='הלל הזקן 14' WHERE id=1")
            A.db().execute("INSERT INTO tasks(site_id, public_no, topic, text, status) VALUES (1,'Z1','חשמל','בדיקה','בטיפול')")
            A.db().commit()
        html = self.login("sup@test").get("/sites/1").get_data(as_text=True)
        self.assertIn("התפלגות סטטוסים", html)
        self.assertIn("אחוז סגירה", html)
        self.assertIn("site-hilel14.jpg", html)
        self.assertIn('class="pill s-prog"', html)
        self.assertEqual(self.login("sup@test").get("/static/site-hilel14.jpg").status_code, 200)


class AdminAndDeleteTests(AppTestCase):
    def _site_with_task(self):
        with A.app.app_context():
            A.db().execute("INSERT INTO tasks(site_id, public_no, topic, text, status) VALUES (1,'X1','חשמל','בדיקה','פתוחה')")
            tid = A.db().execute("SELECT id FROM tasks WHERE public_no='X1'").fetchone()[0]
            A.db().execute("INSERT INTO task_status_log(task_id, status, handled_by, changed_by, changed_at) VALUES (?,?,?,?,?)",
                           (tid, "בטיפול", "", "", ""))
            A.db().commit()
        return tid

    def test_general_admin_screen_is_for_administrators(self):
        self.assertEqual(self.login("owner@test").get("/admin").status_code, 403)
        self.assertEqual(self.login("admin@test").get("/admin").status_code, 200)

    def test_site_deletion_waits_for_approval_link(self):
        tid = self._site_with_task()
        self.login("admin@test").post("/admin/sites/1/delete")
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM sites WHERE id=1"), 1)
        mail = self.db_val("SELECT o.body FROM outbox o WHERE o.mode='מחיקת אתר' ORDER BY o.id DESC LIMIT 1")
        self.assertEqual(self.db_val("SELECT c.email FROM outbox o JOIN contacts c ON c.id=o.contact_id WHERE o.mode='מחיקת אתר' ORDER BY o.id DESC LIMIT 1"),
                         A.DEV_EMAIL)
        token = re.search(r"/admin/approve-delete/([\w-]+)", mail).group(1)
        anon = A.app.test_client()
        self.assertEqual(anon.get(f"/admin/approve-delete/{token}").status_code, 200)
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM sites WHERE id=1"), 1)
        anon.post(f"/admin/approve-delete/{token}")
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM sites WHERE id=1"), 0)
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM tasks WHERE id=?", (tid,)), 0)
        self.assertEqual(self.db_val("SELECT COUNT(*) FROM task_status_log WHERE task_id=?", (tid,)), 0)

    def test_existing_login_users_become_administrators_once(self):
        with A.app.app_context():
            A.db().execute("INSERT INTO contacts(email, first_name, last_name, password_hash, is_admin) VALUES ('new@test','נ','',?,0)",
                           (generate_password_hash("pw"),))
            A.db().execute("DELETE FROM app_settings WHERE key='all_users_admin'")
            A.db().commit()
        A.init_db()
        self.assertEqual(self.db_val("SELECT is_admin FROM contacts WHERE email='new@test'"), 1)

    def test_protocol_screens_show_the_six_steps(self):
        html = self.login("admin@test").get("/sites/1/meetings/new").get_data(as_text=True)
        self.assertIn("עדכון כלל המשתמשים", html)
        self.assertIn('class="step active"', html)


class ProtocolDatesTests(AppTestCase):
    def test_publishing_records_protocol_date_and_publish_time(self):
        today = __import__("datetime").date.today()
        row = {"number": "1", "topic": "חשמל", "task": "משימה לחודש הנוכחי", "assignee": "",
               "target": "", "meeting_date": today.strftime("%d.%m.%y")}
        with A.app.app_context():
            A.db().execute("INSERT INTO meetings(site_id, kind, staged_rows, status) VALUES (1,'ביצוע',?, 'draft')",
                           (json.dumps([row], ensure_ascii=False),))
            A.db().commit()
            A.publish_meeting(A.db().execute("SELECT MAX(id) FROM meetings").fetchone()[0])
        self.assertEqual(self.db_val("SELECT meeting_date FROM meetings ORDER BY id DESC LIMIT 1"), today.isoformat())
        self.assertIsNotNone(self.db_val("SELECT published_at FROM meetings ORDER BY id DESC LIMIT 1"))
        html = self.login("sup@test").get("/sites/1/protocols?month=current").get_data(as_text=True)
        self.assertIn(today.isoformat(), html)
        self.assertEqual(self.login("sup@test").get("/sites/1/protocols?sort=published").status_code, 200)
        self.assertEqual(self.login("owner@test").get("/sites/1/protocols").status_code, 403)

    def test_uploaded_but_unpublished_protocols_are_listed_as_drafts(self):
        with A.app.app_context():
            A.db().execute("INSERT INTO meetings(site_id, kind, source_name, staged_rows, status) VALUES (1,'שבועית','merged.pdf','[]','draft')")
            A.db().commit()
        html = self.login("sup@test").get("/sites/1/protocols").get_data(as_text=True)
        self.assertIn("הועלו ולא פורסמו", html)
        self.assertIn("merged.pdf", html)
        self.assertIn("לא זוהו שורות", html)

    def test_existing_records_get_dates_from_their_protocol_rows(self):
        row = {"number": "1", "topic": "חשמל", "task": "ישן", "assignee": "", "target": "", "meeting_date": "10.11.26"}
        with A.app.app_context():
            A.db().execute("INSERT INTO meetings(site_id, kind, staged_rows, status) VALUES (1,'ביצוע',?, 'published')",
                           (json.dumps([row], ensure_ascii=False),))
            mid = A.db().execute("SELECT MAX(id) FROM meetings").fetchone()[0]
            A.db().execute("INSERT INTO tasks(site_id, public_no, topic, text, status, last_meeting_id) VALUES (1,'OLD','חשמל','ישן','פתוחה',?)", (mid,))
            A.db().commit()
        A.init_db()
        self.assertEqual(self.db_val("SELECT created_at FROM tasks WHERE public_no='OLD'"), "2026-11-10")
        self.assertEqual(self.db_val("SELECT published_at FROM tasks WHERE public_no='OLD'"), "2026-11-10")
        self.assertEqual(self.db_val("SELECT meeting_date FROM meetings WHERE id=?", (mid,)), "2026-11-10")


class DashboardLinksTests(AppTestCase):
    def test_kpi_cards_open_the_matching_task_lists(self):
        with A.app.app_context():
            A.db().execute("INSERT INTO tasks(site_id, public_no, topic, text, status, created_at) VALUES (1,'L1','חשמל','ישן','פתוחה','2020-01-01')")
            A.db().execute("INSERT INTO tasks(site_id, public_no, topic, text, status) VALUES (1,'L2','חשמל','ממתינה','ממתינה לאישור')")
            A.db().commit()
        c = self.login("sup@test")
        html = c.get("/sites/1").get_data(as_text=True)
        for view in ("open", "awaiting", "aged", "repeat2", "done"):
            self.assertIn(f"view={view}", html)
        aged = c.get("/sites/1/search?view=aged").get_data(as_text=True)
        self.assertIn(">L1<", aged)
        self.assertNotIn(">L2<", aged)
        awaiting = c.get("/sites/1/search?view=awaiting").get_data(as_text=True)
        self.assertIn(">L2<", awaiting)
        self.assertNotIn(">L1<", awaiting)


class InitialIdentifiesPersonTests(AppTestCase):
    def setUp(self):
        super().setUp()
        with A.app.app_context():
            d = A.db()
            for email, last in (("ys@x.test", "ישראלי"), ("eg@x.test", "בניה ירוקה")):
                cur = d.execute("INSERT INTO contacts(email, first_name, last_name) VALUES (?,'ליאור',?)", (email, last))
                d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'יועצים')", (cur.lastrowid,))
            d.commit()

    def test_first_name_and_initial_link_to_the_only_matching_person(self):
        with A.app.app_context():
            d = A.db()
            d.execute("INSERT INTO tasks(site_id, public_no, text, assignee_text) VALUES (1,'9','x','ליאור י')")
            tid = d.execute("SELECT id FROM tasks WHERE public_no='9'").fetchone()[0]
            d.commit()
            A.suggest_assignees(tid)
            rows = [tuple(r) for r in d.execute(
                "SELECT c.last_name, a.confirmed FROM task_assignments a JOIN contacts c ON c.id=a.contact_id "
                "WHERE a.task_id=?", (tid,)).fetchall()]
        self.assertEqual(rows, [("ישראלי", 1)])


class TriageTests(AppTestCase):
    def setUp(self):
        super().setUp()
        with A.app.app_context():
            A.db().execute("INSERT INTO tasks(site_id, public_no, text, assignee_text) VALUES (1,'T1','לבדוק','')")
            A.db().commit()
            self.tid = A.db().execute("SELECT id FROM tasks WHERE public_no='T1'").fetchone()[0]

    def test_mark_for_information_only(self):
        self.login("sup@test").post(f"/tasks/{self.tid}/triage", data={"action": "info"})
        self.assertEqual(self.db_val("SELECT status FROM tasks WHERE id=?", (self.tid,)), "לידיעה")

    def test_delete_task(self):
        self.login("sup@test").post(f"/tasks/{self.tid}/triage", data={"action": "delete"})
        self.assertIsNone(self.db_val("SELECT id FROM tasks WHERE id=?", (self.tid,)))

    def test_owner_cannot_triage(self):
        r = self.login("owner@test").post(f"/tasks/{self.tid}/triage", data={"action": "delete"})
        self.assertEqual(r.status_code, 403)
        self.assertIsNotNone(self.db_val("SELECT id FROM tasks WHERE id=?", (self.tid,)))


class ManualAssignTests(AppTestCase):
    def setUp(self):
        super().setUp()
        with A.app.app_context():
            A.db().execute("INSERT INTO tasks(site_id, public_no, text, assignee_text) VALUES (1,'M1','לבדוק','')")
            A.db().commit()
            self.tid = A.db().execute("SELECT id FROM tasks WHERE public_no='M1'").fetchone()[0]
            self.cid = A.db().execute("SELECT id FROM contacts WHERE email='cons@test'").fetchone()[0]

    def test_supervisor_assigns_from_list_and_it_becomes_responsible(self):
        self.login("sup@test").post(f"/tasks/{self.tid}/assign", data={"contact_id": str(self.cid)})
        row = self.db_one("SELECT confirmed, is_responsible FROM task_assignments WHERE task_id=? AND contact_id=?",
                          (self.tid, self.cid))
        self.assertEqual((row["confirmed"], row["is_responsible"]), (1, 1))

    def test_only_members_of_the_site_can_be_assigned(self):
        with A.app.app_context():
            outsider = A.db().execute("INSERT INTO contacts(email, first_name, last_name) VALUES ('out@x.test','זר','')").lastrowid
            A.db().commit()
        r = self.login("sup@test").post(f"/tasks/{self.tid}/assign", data={"contact_id": str(outsider)})
        self.assertEqual(r.status_code, 400)

    def test_warnings_page_offers_contact_list(self):
        html = self.login("sup@test").get("/sites/1/warnings").get_data(as_text=True)
        self.assertIn('name="contact_id"', html)
        self.assertIn("שייך", html)


class ChoiceReplacesSuggestionsTests(AppTestCase):
    def test_manual_choice_removes_other_pending_suggestions_and_they_do_not_return(self):
        with A.app.app_context():
            d = A.db()
            for email, last in (("ys2@x.test", "ישראלי"), ("eg2@x.test", "בניה ירוקה")):
                cur = d.execute("INSERT INTO contacts(email, first_name, last_name) VALUES (?,'ליאור',?)", (email, last))
                d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'יועצים')", (cur.lastrowid,))
            d.execute("INSERT INTO tasks(site_id, public_no, text, assignee_text) VALUES (1,'C1','x','ליאור')")
            d.commit()
            tid = d.execute("SELECT id FROM tasks WHERE public_no='C1'").fetchone()[0]
            A.suggest_assignees(tid)
            keep = d.execute("SELECT id FROM contacts WHERE email='eg2@x.test'").fetchone()[0]
            d.commit()
        self.login("sup@test").post(f"/tasks/{tid}/assign", data={"contact_id": str(keep)})
        self.login("sup@test").get("/sites/1/warnings")  # runs the re-check again
        with A.app.app_context():
            rows = [tuple(r) for r in A.db().execute(
                "SELECT contact_id, confirmed FROM task_assignments WHERE task_id=?", (tid,)).fetchall()]
        self.assertEqual(rows, [(keep, 1)])


class TaskSearchTests(AppTestCase):
    def setUp(self):
        super().setUp()
        with A.app.app_context():
            d = A.db()
            d.execute("INSERT INTO tasks(site_id, public_no, text, assignee_text, status) VALUES (1,'S1','דיפון מערבי של המבנה','ליאור י','פתוחה')")
            d.execute("INSERT INTO tasks(site_id, public_no, text, assignee_text, status) VALUES (1,'S2','תיקון גג','גיא פז','בוצע')")
            d.commit()

    def search(self, q):
        return self.login("owner@test").get("/sites/1/search", query_string={"q": q}).get_data(as_text=True)

    def test_single_word_finds_task(self):
        self.assertIn(">S1<", self.search("דיפון"))
        self.assertNotIn(">S2<", self.search("דיפון"))

    def test_quoted_phrase_must_match_exactly(self):
        self.assertIn(">S1<", self.search('"דיפון מערבי"'))
        self.assertIn("אין תוצאות", self.search('"מערבי דיפון"'))

    def test_several_words_must_all_appear(self):
        self.assertIn(">S2<", self.search("תיקון גיא"))
        self.assertIn("אין תוצאות", self.search("תיקון דיפון"))

    def test_status_is_searchable_and_matches_are_highlighted(self):
        html = self.search("בוצע")
        self.assertIn(">S2<", html)
        self.assertIn("<mark>", html)


class PrintMarkTests(AppTestCase):
    def setUp(self):
        super().setUp()
        with A.app.app_context():
            d = A.db()
            d.execute("INSERT INTO tasks(site_id, public_no, text, assignee_text) VALUES (1,'P1','להדפסה','')")
            d.execute("INSERT INTO tasks(site_id, public_no, text, assignee_text) VALUES (1,'P2','לא להדפסה','')")
            d.commit()
            self.p1 = d.execute("SELECT id FROM tasks WHERE public_no='P1'").fetchone()[0]

    def test_mark_toggles_and_print_page_lists_only_marked(self):
        sup = self.login("sup@test")
        sup.post(f"/tasks/{self.p1}/print-flag", data={"next": "/sites/1/print"})
        html = sup.get("/sites/1/print").get_data(as_text=True)
        self.assertIn(">P1<", html)
        self.assertNotIn(">P2<", html)
        sup.post(f"/tasks/{self.p1}/print-flag", data={"next": "/sites/1/print"})
        self.assertNotIn(">P1<", sup.get("/sites/1/print").get_data(as_text=True))

    def test_clear_removes_all_marks(self):
        sup = self.login("sup@test")
        sup.post(f"/tasks/{self.p1}/print-flag", data={"next": "/"})
        sup.post("/sites/1/print/clear")
        self.assertEqual(self.db_val("SELECT print_flag FROM tasks WHERE id=?", (self.p1,)), 0)

    def test_redirect_only_to_local_path(self):
        self.assertEqual(A._safe_next("//evil.example", "/x"), "/x")
        self.assertEqual(A._safe_next("/sites/1", "/x"), "/sites/1")

    def test_owner_cannot_mark(self):
        self.assertEqual(self.login("owner@test").post(f"/tasks/{self.p1}/print-flag").status_code, 403)


class FilterTests(AppTestCase):
    def setUp(self):
        super().setUp()
        with A.app.app_context():
            d = A.db()
            cur = d.execute("INSERT INTO contacts(email, first_name, last_name) VALUES ('fl@x.test','נועה','ברק')")
            d.execute("INSERT INTO memberships(site_id, contact_id, site_role) VALUES (1,?,'יועצים')", (cur.lastrowid,))
            self.noa = cur.lastrowid
            d.execute("INSERT INTO tasks(site_id, public_no, topic, text, status) VALUES (1,'F1','חשמל','תיקון לוח','פתוחה')")
            d.execute("INSERT INTO tasks(site_id, public_no, topic, text, status) VALUES (1,'F2','חשמל','סגירה','בוצע')")
            d.execute("INSERT INTO tasks(site_id, public_no, topic, text, status) VALUES (1,'F3','אינסטלציה','צנרת','פתוחה')")
            t1 = d.execute("SELECT id FROM tasks WHERE public_no='F1'").fetchone()[0]
            d.execute("INSERT INTO task_assignments(task_id, contact_id, raw_name, confirmed) VALUES (?,?, '', 1)",
                      (t1, self.noa))
            d.execute("INSERT INTO decisions(site_id, topic, text, decided_on, source, created) "
                      "VALUES (1,'חשמל','לאשר לוח חדש','02.12.25','פרוטוקול','x')")
            d.commit()

    def get(self, **params):
        return self.login("owner@test").get("/sites/1/search", query_string=params).get_data(as_text=True)

    def test_open_view_hides_done_tasks(self):
        html = self.get(view="open")
        self.assertIn(">F1<", html)
        self.assertNotIn(">F2<", html)

    def test_topic_and_status_combine(self):
        html = self.get(topic="חשמל", status="בוצע")
        self.assertIn(">F2<", html)
        self.assertNotIn(">F1<", html)

    def test_contact_filter_uses_assignments(self):
        html = self.get(contact=str(self.noa))
        self.assertIn(">F1<", html)
        self.assertNotIn(">F3<", html)

    def test_decisions_kind_lists_decisions_and_filters_by_topic(self):
        self.assertIn("לאשר לוח חדש", self.get(kind="החלטות"))
        self.assertIn("אין החלטות תואמות", self.get(kind="החלטות", topic="אינסטלציה"))


if __name__ == "__main__":
    unittest.main()
