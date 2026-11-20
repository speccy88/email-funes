"""Regression tests for export snapshots, identities, and failed staging."""
import json
from pathlib import Path
import sys
import tempfile
import unittest

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

    def tearDown(self):
        self.db.close()
        self.tmp.cleanup()

    def test_existing_snapshot_is_preserved(self):
        export.map_new_messages(self.db)
        target = self.root / 'first'
        self.assertEqual(export.export(self.db, target, per_file=1), 3)
        original = {p.name: p.read_bytes() for p in target.iterdir()}
        with self.assertRaisesRegex(ValueError, 'new or empty'):
            export.export(self.db, target, per_file=20)
        self.assertEqual(original, {p.name: p.read_bytes() for p in target.iterdir()})

    def test_failed_export_exposes_no_partial_snapshot(self):
        export.map_new_messages(self.db)
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
        added, missing = export.map_new_messages(self.db)
        self.assertEqual((added, missing), (2, 1))
        turns = list(export.turns(self.db))
        self.assertEqual(len(turns), 2)
        self.assertTrue(all(t['seq'] == 0 for t in turns))

    def test_sequence_gap_fails_instead_of_publishing_invalid_session(self):
        export.map_new_messages(self.db)
        self.db.execute('UPDATE funes_map SET seq=5 WHERE message_id=2')
        self.db.commit()
        with self.assertRaisesRegex(ValueError, 'Non-dense'):
            export.export(self.db, self.root / 'gap')

    def test_repeated_snapshots_preserve_turn_identities(self):
        export.map_new_messages(self.db)
        first = list(export.turns(self.db))
        self.assertEqual(export.map_new_messages(self.db), (0, 0))
        self.assertEqual(first, list(export.turns(self.db)))
        for turn in first:
            self.assertEqual(set(turn), {'format', 'session_id', 'turn_uuid', 'seq', 'ts', 'role', 'harness', 'blocks'})
            self.assertNotIn(':', turn['session_id'])
            self.assertNotIn(':', turn['turn_uuid'])
            self.assertTrue(turn['ts'].endswith('Z'))


if __name__ == '__main__':
    unittest.main()
