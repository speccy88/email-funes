"""Regression tests for export snapshots, identities, and failed staging."""
import json
from contextlib import redirect_stderr, redirect_stdout
import io
import os
from pathlib import Path
import sqlite3
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
sys.path.insert(0, str(ROOT / 'examples'))

import email_funes_common as common
import sqlite_to_funes as export
from create_demo_db import demo_records


class ExportSafetyTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.db = common.open_database(self.root / 'demo.sqlite')
        for record in demo_records():
            common.store_message(self.db, record)
        self.db.commit()
        export.ensure_map(self.db)
        self.legacy = export.resolve_policy(self.db, include_all=True)

    def tearDown(self):
        self.db.close()
        self.tmp.cleanup()

    def test_existing_snapshot_is_preserved(self):
        export.map_new_messages(self.db, self.legacy)
        target = self.root / 'first'
        self.assertEqual(export.export(self.db, target, per_file=1), 3)
        original = {p.name: p.read_bytes() for p in target.iterdir()}
        with self.assertRaisesRegex(ValueError, 'new or empty'):
            export.export(self.db, target, per_file=20)
        self.assertEqual(original, {p.name: p.read_bytes() for p in target.iterdir()})

    def test_failed_export_exposes_no_partial_snapshot(self):
        export.map_new_messages(self.db, self.legacy)
        self.db.execute("UPDATE messages SET attachment_names_json='invalid json' WHERE id=2")
        self.db.commit()
        target = self.root / 'failed'
        with self.assertRaises(json.JSONDecodeError):
            export.export(self.db, target, per_file=1)
        self.assertFalse(target.exists())
        self.assertEqual(list(self.root.glob('.email-funes-export-*')), [])

    def test_missing_dates_are_not_invented(self):
        self.db.execute('UPDATE messages SET sent_utc=NULL,received_utc=NULL WHERE id=1')
        self.db.commit()
        added, missing = export.map_new_messages(self.db, self.legacy)
        self.assertEqual((added, missing), (2, 1))
        turns = list(export.turns(self.db))
        self.assertEqual(len(turns), 2)
        self.assertTrue(all(t['seq'] == 0 for t in turns))

    def test_sequence_gap_fails_instead_of_publishing_invalid_session(self):
        export.map_new_messages(self.db, self.legacy)
        self.db.execute('UPDATE funes_map SET seq=5 WHERE message_id=2')
        self.db.commit()
        with self.assertRaisesRegex(ValueError, 'Non-dense'):
            export.export(self.db, self.root / 'gap')

    def test_repeated_snapshots_preserve_turn_identities(self):
        export.map_new_messages(self.db, self.legacy)
        first = list(export.turns(self.db))
        self.assertEqual(export.map_new_messages(self.db), (0, 0))
        self.assertEqual(first, list(export.turns(self.db)))
        for turn in first:
            self.assertEqual(set(turn), {'format', 'session_id', 'turn_uuid', 'seq', 'ts', 'role', 'harness', 'blocks'})
            self.assertNotIn(':', turn['session_id'])
            self.assertNotIn(':', turn['turn_uuid'])
            self.assertTrue(turn['ts'].endswith('Z'))

    def run_cli(self, out):
        diagnostics = io.StringIO()
        with redirect_stdout(diagnostics), redirect_stderr(diagnostics):
            rc = export.main(["--db", str(self.root / "demo.sqlite"), "--include-all", "--out", str(out)])
        return rc, diagnostics.getvalue()

    def test_render_failure_rolls_back_new_mapping_and_redacts_invalid_fields(self):
        self.db.execute("UPDATE messages SET sent_utc=? WHERE id=1", ("private\x1b[31m@invalid.exampleZ",))
        self.db.commit()
        target = self.root / "sensitive-name"
        result, diagnostics = self.run_cli(target)
        self.assertNotEqual(result, 0)
        for private in ("sensitive-name", "invalid.example", "\x1b"):
            self.assertNotIn(private, diagnostics)
        self.assertFalse(target.exists())
        self.assertFalse(export.table_exists(self.db, "funes_export_policy"))
        self.assertEqual(self.db.execute("SELECT COUNT(*) FROM funes_map").fetchone()[0], 0)

    def test_reserved_database_and_sidecar_paths(self):
        path = self.root / "demo.sqlite"
        for target in (path, Path(str(path) + "-wal"), Path(str(path) + "-journal") / "out"):
            with self.subTest(target=target), self.assertRaisesRegex(ValueError, "collides"):
                export.check_destination(target, path)
        hardlink = self.root / "hardlink"
        os.link(path, hardlink)
        self.assertNotEqual(self.run_cli(hardlink)[0], 0)
        self.assertTrue(path.samefile(hardlink))

    def test_unrelated_input_remains_unchanged(self):
        other = self.root / "unrelated.sqlite"
        connection = sqlite3.connect(other)
        connection.execute("CREATE TABLE unrelated (id INTEGER)")
        connection.close()
        plain = self.root / "another-file"
        plain.write_bytes(b"Synthetic private content")
        empty = self.root / "empty.sqlite"
        empty.touch()
        schema_empty = self.root / "empty-schema.sqlite"
        connection = sqlite3.connect(schema_empty)
        connection.execute("PRAGMA user_version=9")
        connection.close()
        for path in (other, plain, empty, schema_empty):
            if os.name == "posix":
                path.chmod(0o644)
            mode = path.stat().st_mode
            before = path.read_bytes()
            with self.subTest(path=path), redirect_stderr(io.StringIO()):
                self.assertNotEqual(export.main(["--db", str(path), "--out", str(self.root / "out")]), 0)
                self.assertEqual(path.read_bytes(), before)
                self.assertEqual(path.stat().st_mode, mode)
                self.assertFalse((self.root / "out").exists())

    def test_symlink_output_is_rejected_even_if_empty_or_dangling(self):
        target = self.root / "empty"
        target.mkdir()
        out = self.root / "link"
        try:
            out.symlink_to(target, target_is_directory=True)
        except OSError:
            self.skipTest("Directory symlinks require host support")
        for _ in range(2):
            self.assertNotEqual(self.run_cli(out)[0], 0)
            self.assertTrue(out.is_symlink())
            if target.exists():
                target.rmdir()

    def test_concurrent_output_is_preserved(self):
        export.map_new_messages(self.db, self.legacy)
        out = self.root / "destination"
        original_turns = export.turns

        def concurrent(db, policy):
            yield from original_turns(db, policy)
            out.mkdir()
            (out / "keep").write_text("other output", encoding="utf-8")

        with patch.object(export, "turns", concurrent), self.assertRaisesRegex(ValueError, "new or empty"):
            export.export(self.db, out)
        self.assertEqual([p.name for p in out.iterdir()], ["keep"])
        self.assertEqual(list(self.root.glob(".email-funes-export-*")), [])

    def test_existing_empty_directory_survives_render_failure_and_accepts_retry(self):
        out = self.root / "empty"
        out.mkdir()
        with patch.object(export, "new_private_text", side_effect=OSError("sensitive detail")):
            result, diagnostics = self.run_cli(out)
        self.assertNotEqual(result, 0)
        self.assertNotIn("sensitive detail", diagnostics)
        self.assertTrue(out.is_dir())
        self.assertEqual(list(out.iterdir()), [])
        self.assertFalse(export.table_exists(self.db, "funes_export_policy"))
        self.assertEqual(list(self.root.glob(".email-funes-export-*")), [])
        self.assertEqual(self.run_cli(out)[0], 0)
        self.assertEqual(len(list(out.glob("*.funes.jsonl"))), 1)

    def test_rename_failure_after_commit_keeps_durable_ids_for_retry(self):
        out = self.root / "new-snapshot"
        with patch.object(Path, "rename", side_effect=PermissionError("private location")):
            result, diagnostics = self.run_cli(out)
        self.assertNotEqual(result, 0)
        self.assertNotIn("private location", diagnostics)
        self.assertFalse(out.exists())
        self.assertEqual(list(self.root.glob(".email-funes-export-*")), [])
        expected = list(export.turns(self.db))
        self.assertEqual(len(expected), 3)
        self.assertEqual(self.run_cli(out)[0], 0)
        lines = next(out.glob("*.funes.jsonl")).read_text(encoding="utf-8").splitlines()
        self.assertEqual([json.loads(line) for line in lines], expected)

    @unittest.skipUnless(os.name == "posix", "POSIX mode bits")
    def test_private_file_and_directory_modes_under_permissive_umask(self):
        import stat
        export.map_new_messages(self.db, self.legacy)
        existing = self.root / "public-parent"
        existing.mkdir(mode=0o755)
        before = existing.stat().st_mode
        target = existing / "private" / "snapshot"
        previous = os.umask(0)
        try:
            export.export(self.db, target)
        finally:
            os.umask(previous)
        self.assertEqual(existing.stat().st_mode, before)
        for private in (target.parent, target):
            self.assertEqual(stat.S_IMODE(private.stat().st_mode), 0o700)
        for file in target.glob("*.jsonl"):
            self.assertEqual(stat.S_IMODE(file.stat().st_mode), 0o600)


if __name__ == '__main__':
    unittest.main()
