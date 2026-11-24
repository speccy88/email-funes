"""Synthetic regression tests for recovered metadata and stable export choices."""
from contextlib import redirect_stderr, redirect_stdout
import hashlib
import io
import json
from pathlib import Path
import sqlite3
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "examples"))

import email_funes_common as common
import sqlite_to_funes as export
from create_demo_db import demo_message


class EnrichedExportTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.path = self.root / "mail.sqlite"
        self.db = common.open_database(self.path)
        for number in range(1, 6):
            common.store_message(self.db, demo_message(
                number, f"Synthetic subject {number}", f"Synthetic body {number}",
                f"2026-11-0{number}T14:00:00Z", "sender@example.com",
                "legacy-to@example.com", "shared-session",
            ))
        self.db.execute("UPDATE messages SET cc_list='legacy-cc@example.com'")
        self.db.commit()

    def tearDown(self):
        self.db.close()
        self.tmp.cleanup()

    def enrich(self):
        self.db.executescript("""
            ALTER TABLE messages ADD COLUMN message_category TEXT NOT NULL DEFAULT '';
            CREATE TABLE recovery_recipients (
                message_id INTEGER NOT NULL REFERENCES messages(id),
                ordinal INTEGER NOT NULL,
                recipient_type TEXT NOT NULL,
                display_name TEXT NOT NULL,
                address_type TEXT NOT NULL,
                address_raw TEXT NOT NULL,
                smtp_address TEXT NOT NULL,
                PRIMARY KEY(message_id,ordinal)
            );
            CREATE TABLE recovery_metadata (
                message_id INTEGER PRIMARY KEY REFERENCES messages(id),
                display_bcc TEXT NOT NULL DEFAULT ''
            );
        """)
        categories = ["Email / IPM.Note", "Meeting request", "Email / IPM.Note", "Appointment", ""]
        self.db.executemany("UPDATE messages SET message_category=? WHERE id=?",
                            [(value, number) for number, value in enumerate(categories, 1)])
        self.db.commit()

    def recipient(self, message_id, ordinal, kind, name, smtp, raw=""):
        self.db.execute("INSERT INTO recovery_recipients VALUES(?,?,?,?,?,?,?)",
                        (message_id, ordinal, kind, name, "SMTP", raw, smtp))
        self.db.commit()

    def map(self, **options):
        policy = export.resolve_policy(self.db, **options)
        export.ensure_map(self.db)
        export.map_new_messages(self.db, policy)
        return policy

    def test_enriched_default_filters_before_assigning_dense_sequence(self):
        self.enrich()
        policy = self.map()
        self.assertEqual(policy.selection, "email")
        self.assertEqual(self.db.execute("SELECT message_id,seq FROM funes_map ORDER BY seq").fetchall(),
                         [(1, 0), (3, 1)])
        records = list(export.turns(self.db, policy))
        self.assertEqual(len(records), 2)
        self.assertEqual([record["seq"] for record in records], [0, 1])
        self.assertIn("Outlook item category: Email / IPM.Note", records[0]["blocks"][0]["text"])

    def test_meeting_and_all_selection_are_explicit(self):
        self.enrich()
        meetings = export.summarize(self.db, export.resolve_policy(self.db, include_meetings=True))
        all_items = export.summarize(self.db, export.resolve_policy(self.db, include_all=True))
        self.assertEqual(meetings["selected"], 3)
        self.assertEqual(all_items["selected"], 5)
        self.assertEqual(all_items["categories"]["Unclassified"], 1)

    def test_normalized_recipients_and_per_type_fallback(self):
        self.enrich()
        self.recipient(1, 0, "to", "Engineering", "current@example.com", "/o=Example/ou=Recipients/cn=current")
        self.recipient(1, 1, "bcc", "Confidential", "bcc@example.com")
        self.recipient(3, 0, "cc", "Planning", "planning@example.com")
        policy = self.map()
        records = list(export.turns(self.db, policy))
        first, second = [record["blocks"][0]["text"] for record in records]
        self.assertIn("To: Engineering <current@example.com>", first)
        self.assertIn("Cc: legacy-cc@example.com", first)
        self.assertNotIn("/o=Example", first)
        self.assertNotIn("bcc@example.com", first)
        self.assertNotIn("Bcc:", first)
        self.assertIn("To: legacy-to@example.com", second)
        self.assertIn("Cc: Planning <planning@example.com>", second)

    def test_bcc_opt_in_includes_typed_rows_and_display_fallback(self):
        self.enrich()
        self.recipient(1, 0, "bcc", "Reviewer", "reviewer@example.com")
        self.db.execute("INSERT INTO recovery_metadata VALUES(3,'fallback-bcc@example.com')")
        self.db.commit()
        policy = self.map(include_bcc=True)
        text = "\n".join(record["blocks"][0]["text"] for record in export.turns(self.db, policy))
        self.assertIn("Bcc: Reviewer <reviewer@example.com>", text)
        self.assertIn("Bcc: fallback-bcc@example.com", text)

    def test_numeric_recipient_types_and_unresolved_display_name(self):
        self.enrich()
        self.recipient(1, 0, "1", "No address recovered", "")
        self.recipient(1, 1, "2", "Cc person", "cc@example.com")
        policy = self.map()
        text = next(export.turns(self.db, policy))["blocks"][0]["text"]
        self.assertIn("To: No address recovered", text)
        self.assertIn("Cc: Cc person <cc@example.com>", text)

    def test_policy_change_is_rejected_without_mutating_policy_or_mapping(self):
        self.enrich()
        policy = self.map()
        before_policy = self.db.execute("SELECT policy_json FROM funes_export_policy").fetchone()[0]
        before_mapping = self.db.execute("SELECT * FROM funes_map ORDER BY message_id").fetchall()
        for options in ({"include_meetings": True}, {"include_bcc": True}, {"include_all": True}):
            with self.assertRaisesRegex(ValueError, "Export policy differs"):
                export.map_new_messages(self.db, export.resolve_policy(self.db, **options))
        self.assertEqual(self.db.execute("SELECT policy_json FROM funes_export_policy").fetchone()[0], before_policy)
        self.assertEqual(self.db.execute("SELECT * FROM funes_map ORDER BY message_id").fetchall(), before_mapping)
        self.assertEqual(export.map_new_messages(self.db, policy), (0, 0))

    def test_reclassification_cannot_remove_previously_mapped_messages(self):
        self.enrich()
        policy = self.map()
        self.db.execute("UPDATE messages SET message_category='Appointment' WHERE id=1")
        self.db.commit()
        with self.assertRaisesRegex(ValueError, "missing or excluded"):
            list(export.turns(self.db, policy))

    def test_legacy_database_and_policyless_mapping_stay_compatible(self):
        policy = self.map(include_all=True)
        self.assertEqual(policy, export.ExportPolicy())
        original = list(export.turns(self.db, policy))
        self.db.execute("DROP TABLE funes_export_policy")
        self.db.commit()
        self.assertEqual(export.map_new_messages(self.db), (0, 0))
        self.assertEqual(list(export.turns(self.db, policy)), original)
        self.assertEqual(len(original), 5)

    def test_enriching_policyless_legacy_mapping_requires_deliberate_rebuild(self):
        self.map(include_all=True)
        self.db.execute("DROP TABLE funes_export_policy")
        self.db.commit()
        self.enrich()
        with self.assertRaisesRegex(ValueError, "legacy rendering"):
            export.map_new_messages(self.db, export.resolve_policy(self.db))
        self.assertFalse(export.table_exists(self.db, "funes_export_policy"))
        self.assertEqual(self.db.execute("SELECT COUNT(*) FROM funes_map").fetchone()[0], 5)

    def test_dry_run_reads_only_and_never_creates_export_or_mapping(self):
        self.enrich()
        self.db.execute("UPDATE messages SET sent_utc=NULL,received_utc=NULL WHERE id=3")
        self.db.commit()
        self.db.close()
        before = hashlib.sha256(self.path.read_bytes()).hexdigest()
        output = io.StringIO()
        with redirect_stdout(output):
            result = export.main(["--db", str(self.path), "--dry-run"])
        self.assertEqual(result, 0)
        self.assertEqual(hashlib.sha256(self.path.read_bytes()).hexdigest(), before)
        self.assertIn("selected: 2; excluded: 3", output.getvalue())
        self.assertIn("missing dates: 1", output.getvalue())
        self.assertNotIn("sender@example.com", output.getvalue())
        self.db = sqlite3.connect(self.path)
        self.assertFalse(export.table_exists(self.db, "funes_map"))
        self.assertFalse(export.table_exists(self.db, "funes_export_policy"))
        self.assertEqual(list(self.root.glob("*.jsonl")), [])

    def test_missing_date_skip_does_not_create_sequence_gap(self):
        self.enrich()
        self.db.execute("UPDATE messages SET sent_utc=NULL,received_utc=NULL WHERE id=1")
        self.db.commit()
        policy = self.map()
        records = list(export.turns(self.db, policy))
        self.assertEqual([record["seq"] for record in records], [0])

    def test_legacy_cli_notice_and_enriched_snapshot_format(self):
        with redirect_stderr(io.StringIO()) as error:
            for args in (["--dry-run"], ["--out", str(self.root / "blocked")]):
                self.assertEqual(export.main(["--db", str(self.path), *args]), 2)
        self.assertIn("--include-all", error.getvalue())
        self.assertFalse((self.root / "blocked").exists())
        self.assertFalse(export.table_exists(self.db, "funes_map"))
        with redirect_stdout(io.StringIO()) as output:
            self.assertEqual(export.main(["--db", str(self.path), "--dry-run", "--include-all"]), 0)
        self.assertIn("Legacy database", output.getvalue())
        self.enrich()
        policy = self.map()
        target = self.root / "snapshot"
        self.assertEqual(export.export(self.db, target, per_file=1, policy=policy), 2)
        records = [json.loads(path.read_text(encoding="utf-8")) for path in sorted(target.glob("*.funes.jsonl"))]
        self.assertEqual([record["seq"] for record in records], [0, 1])
        for record in records:
            self.assertEqual(set(record), {"format", "session_id", "turn_uuid", "seq", "ts", "role", "harness", "blocks"})
            self.assertEqual(record["role"], "member")
            self.assertNotIn("bcc@example.com", record["blocks"][0]["text"])

    def test_saved_selection_reused_and_saved_bcc_requires_confirmation(self):
        self.enrich()
        policy = self.map(include_meetings=True, include_bcc=True)
        with self.assertRaisesRegex(ValueError, "Bcc"):
            export.resolve_policy(self.db)
        self.assertEqual(export.resolve_policy(self.db, include_bcc=True), policy)

    def test_schema_change_cannot_silently_switch_a_stored_rendering_policy(self):
        policy = self.map(include_all=True)
        self.enrich()
        with self.assertRaisesRegex(ValueError, "incompatible"):
            export.resolve_policy(self.db)
        self.assertEqual(export.stored_policy_json(self.db), policy.serialize())

    def test_diagnostics_hide_unrecognized_category_and_keep_original_evidence(self):
        self.enrich()
        secret = "Sensitive\x1b[2K custom field"
        self.db.execute("UPDATE messages SET message_category=? WHERE id=1", (secret,))
        self.db.commit()
        with redirect_stdout(io.StringIO()) as output:
            self.assertEqual(export.main(["--db", str(self.path), "--dry-run", "--include-all"]), 0)
        self.assertIn("Other category: 1", output.getvalue())
        self.assertNotIn(secret, output.getvalue())
        self.assertEqual(self.db.execute("SELECT message_category FROM messages WHERE id=1").fetchone()[0], secret)


if __name__ == "__main__":
    unittest.main()
