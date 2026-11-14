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

    def test_reimport_does_not_duplicate_memberships(self):
        path = self.write_xlsx([("יועצים", [["שם פרטי", "דוא\"ל"], ["ענת", "anat@example.com"]])])
        with A.app.app_context():
            A.import_contacts(path, site_id=1, site_role="יועצים")
            A.import_contacts(path, site_id=1, site_role="יועצים")
            n = A.db().execute("SELECT COUNT(*) FROM memberships m JOIN contacts c ON c.id=m.contact_id "
                               "WHERE c.email='anat@example.com' AND m.site_id=1").fetchone()[0]
        self.assertEqual(n, 1)


if __name__ == "__main__":
    unittest.main()
