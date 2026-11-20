"""Regression tests for message retention, batch writes, and snapshot resume."""
import base64
from contextlib import closing, redirect_stderr
from io import StringIO
import os
from pathlib import Path
import sqlite3
import sys
import tempfile
from types import SimpleNamespace
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))

import email_funes_common as common
import pst_to_sqlite as outlook
import pst_to_sqlite_direct as direct
from test_converter import FakeItem
from test_direct import FakeFolder, FakeMessage, FakePFF


class ImportSafetyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name)
        self.db_path = self.path / 'mail.sqlite'
        self.db = common.open_database(self.db_path)
        self.pst = self.path / 'mail.pst'
        self.pst.write_bytes(b'not a real PST: tests inject a fake reader')
        self.args = SimpleNamespace(
            db=self.db_path, limit=0, progress=1000, batch_size=250,
            exclude_junk=True, exclude_deleted=True, include_header_spam=False,
            skip_bulk=False, max_body_chars=0,
        )

    def tearDown(self):
        self.db.close()
        self.temp.cleanup()

    def test_recipients_are_part_of_duplicate_comparison(self):
        first = FakeItem('a', '<same@example.org>', 'Notice', 'Identical body')
        second = FakeItem('b', '<same@example.org>', 'Notice', 'Identical body')
        first.To = 'alice@example.org'
        second.To = 'bob@example.org'
        a, _ = outlook.build_message(first, 'archive.pst', 'Inbox', False, False, 0)
        b, _ = outlook.build_message(second, 'archive.pst', 'Inbox', False, False, 0)
        self.assertNotEqual(a.fingerprint, b.fingerprint)
        self.assertEqual(common.store_message(self.db, a), 'kept')
        self.assertEqual(common.store_message(self.db, b), 'kept')

    def test_truncated_prefix_does_not_collapse_distinct_bodies(self):
        for build in ('outlook', 'direct'):
            with self.subTest(reader=build):
                if build == 'outlook':
                    first = FakeItem('a', '<same@example.org>', 'Notice', 'Prefix: first ending')
                    second = FakeItem('b', '<same@example.org>', 'Notice', 'Prefix: second ending')
                    a, _ = outlook.build_message(first, 'mail.pst', 'Inbox', False, False, 7)
                    b, _ = outlook.build_message(second, 'mail.pst', 'Inbox', False, False, 7)
                else:
                    self.args.max_body_chars = 7
                    a, _ = direct.build_message(FakeMessage(body='Prefix: first ending'), self.pst, 'Inbox', 0, self.args)
                    b, _ = direct.build_message(FakeMessage(body='Prefix: second ending'), self.pst, 'Inbox', 1, self.args)
                self.assertEqual(a.body, b.body)
                self.assertNotEqual(a.fingerprint, b.fingerprint)

    def test_batch_commit_is_controlled_by_caller(self):
        item = FakeItem('a', '<a@example.org>', 'Subject', 'Body')
        message, _ = outlook.build_message(item, 'mail.pst', 'Inbox', False, False, 0)
        common.store_message(self.db, message)
        self.assertTrue(self.db.in_transaction)
        with closing(sqlite3.connect(self.db_path)) as other:
            self.assertEqual(other.execute('SELECT COUNT(*) FROM messages').fetchone()[0], 0)
            self.db.commit()
            self.assertEqual(other.execute('SELECT COUNT(*) FROM messages').fetchone()[0], 1)

    def test_source_guard_accepts_unchanged_snapshot(self):
        first = direct.register_pst_source(self.db, self.pst)
        self.assertEqual(direct.register_pst_source(self.db, self.pst), first)

    def test_source_guard_rejects_changed_size(self):
        direct.register_pst_source(self.db, self.pst)
        self.pst.write_bytes(b'changed snapshot with different size')
        with self.assertRaisesRegex(RuntimeError, 'PST changed'):
            direct.register_pst_source(self.db, self.pst)

    def test_source_guard_rejects_changed_modification_time(self):
        direct.register_pst_source(self.db, self.pst)
        stat = self.pst.stat()
        os.utime(self.pst, ns=(stat.st_atime_ns, stat.st_mtime_ns + 1_000_000_000))
        with self.assertRaisesRegex(RuntimeError, 'PST changed'):
            direct.register_pst_source(self.db, self.pst)

    def test_old_database_baseline_is_explicit(self):
        item = FakeItem('a', '<a@example.org>', 'Subject', 'Body')
        message, _ = outlook.build_message(item, 'mail.pst', 'Inbox', False, False, 0)
        common.store_message(self.db, message)
        self.db.commit()
        output = StringIO()
        with redirect_stderr(output):
            direct.register_pst_source(self.db, self.pst)
        self.assertIn('Earlier imports cannot be verified', output.getvalue())

    def test_header_root_and_reply_share_conversation_key(self):
        root = FakeMessage(headers='Message-ID: <root@example.org>\r\n', index=b'')
        reply = FakeMessage(headers='Message-ID: <reply@example.org>\r\nReferences: <root@example.org>\r\n', index=b'')
        first, _ = direct.build_message(root, self.pst, 'Inbox', 0, self.args)
        second, _ = direct.build_message(reply, self.pst, 'Inbox', 1, self.args)
        self.assertEqual(first.conversation_key, second.conversation_key)

    def test_thread_index_header_used_when_property_missing(self):
        root = bytes(range(22))
        first = FakeMessage(headers='Thread-Index: ' + base64.b64encode(root).decode() + '\r\n', index=b'')
        second = FakeMessage(headers='Thread-Index: ' + base64.b64encode(root + bytes(5)).decode() + '\r\n', index=b'')
        a, _ = direct.build_message(first, self.pst, 'Inbox', 0, self.args)
        b, _ = direct.build_message(second, self.pst, 'Inbox', 1, self.args)
        self.assertTrue(a.conversation_key.startswith('outlook-'))
        self.assertEqual(a.conversation_key, b.conversation_key)

    def test_unreadable_body_is_reported_and_not_stored(self):
        fake = FakeMessage()

        class Unreadable:
            @property
            def plain_text_body(self):
                raise OSError('synthetic read failure')

            def __getattr__(self, name):
                return getattr(fake, name)

        root = FakeFolder('Root', [Unreadable()])
        module = SimpleNamespace(file=lambda: FakePFF(root))
        self.assertEqual(direct.import_pst(self.db, self.pst, self.args, module), 2)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM messages').fetchone()[0], 0)
        self.assertIn('synthetic read failure', self.db_path.with_suffix('.direct.errors.txt').read_text())

    def test_unreadable_folder_count_does_not_become_zero(self):
        class UnreadableFolder:
            @property
            def number_of_sub_messages(self):
                raise OSError('synthetic count failure')

        with self.assertRaisesRegex(OSError, 'synthetic count failure'):
            direct.folder_count(UnreadableFolder(), 'messages')


if __name__ == '__main__':
    unittest.main()
