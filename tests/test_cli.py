"""Public import dispatch and source identities across archives."""

from contextlib import closing, redirect_stderr, redirect_stdout
from datetime import datetime
from io import StringIO
import json
from pathlib import Path
import sqlite3
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from email_funes.cli import main
from test_direct import FakeFolder, FakeMessage, FakePFF
from test_metadata_audit import Folder, Message, Module


class CommandTests(unittest.TestCase):
    def test_mocked_reader_import_audit_recover_preview_export(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pst = root / "test.pst"
            pst.write_bytes(b"synthetic reader fixture")
            db, enriched, out = root / "input.sqlite", root / "enriched.sqlite", root / "turns"
            message = Message()
            message.subject = "Mock inspection"
            message.plain_text_body = "Inspect the coupling."
            message.client_submit_time = datetime(2026, 11, 2)
            module = Module(Folder("Root", folders=[Folder("", [message])]))
            commands = [
                ("import", pst, "--format", "pst", "--db", db),
                ("audit", "--pst", pst, "--db", db, "--out", root / "audit.log"),
                ("recover", "--pst", pst, "--db", db, "--out-db", enriched, "--log", root / "recovery.log"),
                ("stats", "--db", enriched),
                ("export", "--db", enriched, "--dry-run"),
                ("export", "--db", enriched, "--out", out),
            ]
            with patch.dict(sys.modules, {"pypff": module}):
                for command in commands:
                    with self.subTest(command=command[0]), redirect_stdout(StringIO()), redirect_stderr(StringIO()):
                        self.assertEqual(main([str(arg) for arg in command]), 0)
            (jsonl,) = out.glob("*.funes.jsonl")
            turn = json.loads(jsonl.read_text(encoding="utf-8"))
            self.assertEqual(turn["harness"], "outlook-pst")
            self.assertIn("Inspect the coupling.", json.dumps(turn))
            self.assertEqual(pst.read_bytes(), b"synthetic reader fixture")

    def test_same_filename_separate_archives_and_repeat_import(self):
        with tempfile.TemporaryDirectory() as directory, redirect_stdout(StringIO()), redirect_stderr(StringIO()):
            dest = Path(directory) / "collection.sqlite"
            for label in ("first", "second"):
                archive = Path(directory) / label / "mail.pst"
                archive.parent.mkdir()
                archive.write_bytes(label.encode())
                root = FakeFolder("Inbox", [FakeMessage(), FakeMessage(subject=label)])
                module = SimpleNamespace(file=lambda: FakePFF(root))
                command = ["import", str(archive), "--format", "pst", "--db", str(dest)]
                with patch.dict(sys.modules, {"pypff": module}):
                    self.assertEqual(main(command), 0)
                    self.assertEqual(main(command), 0)
            with closing(sqlite3.connect(dest)) as db:
                self.assertEqual(db.execute("SELECT COUNT(*) FROM import_sources").fetchone()[0], 2)
                self.assertEqual(db.execute("SELECT COUNT(*) FROM source_items").fetchone()[0], 4)
                self.assertEqual(db.execute("SELECT COUNT(*) FROM messages").fetchone()[0], 3)

    def test_stats_rejects_missing_archive_without_creating_it(self):
        with tempfile.TemporaryDirectory() as directory, redirect_stderr(StringIO()) as output:
            missing = Path(directory) / "private-filename.sqlite"
            self.assertNotEqual(main(["stats", "--db", str(missing)]), 0)
            self.assertFalse(missing.exists())
            self.assertNotIn(missing.name, output.getvalue())

    def test_recovery_checks_optional_rtf_before_creating_outputs(self):
        with tempfile.TemporaryDirectory() as directory, redirect_stderr(StringIO()) as output:
            root = Path(directory)
            pst, db, out, log = (root / n for n in ("input.pst", "db.sqlite", "out.sqlite", "recovery.log"))
            pst.touch()
            db.touch()
            with patch.dict(sys.modules, {"pypff": object(), "striprtf.striprtf": None}):
                with self.assertRaises(SystemExit) as result:
                    main(["recover", "--pst", str(pst), "--db", str(db), "--out-db", str(out),
                          "--log", str(log), "--recover-rtf"])
            self.assertEqual(result.exception.code, 2)
            self.assertIn("striprtf", output.getvalue())
            self.assertFalse(out.exists())
            self.assertFalse(log.exists())
