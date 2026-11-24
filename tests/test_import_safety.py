"""Regression tests for message retention, batch writes, and snapshot resume."""
import base64
from contextlib import closing, redirect_stderr, redirect_stdout
from io import StringIO
import os
from pathlib import Path
import sqlite3
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))

import email_funes_common as common
import email_funes_files as files
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
                raise OSError('synthetic private failure\x1b[2J')

            def __getattr__(self, name):
                return getattr(fake, name)

        root = FakeFolder('Private diagnostics folder', [Unreadable()])
        module = SimpleNamespace(file=lambda: FakePFF(root))
        console = StringIO()
        with redirect_stdout(console), redirect_stderr(console):
            self.assertEqual(direct.import_pst(self.db, self.pst, self.args, module), 2)
            self.args.private_diagnostics = True
            self.assertEqual(direct.import_pst(self.db, self.pst, self.args, module), 2)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM messages').fetchone()[0], 0)
        self.assertEqual(self.db_path.with_suffix('.direct.errors.txt').read_text(), 'item.read error=OSError\n')
        private = self.db_path.with_suffix('.direct.private.errors.txt').read_text()
        for field in ('synthetic private failure', 'Private diagnostics folder'):
            self.assertIn(field, private)
            self.assertNotIn(field, console.getvalue())
        self.assertNotIn(str(self.pst), console.getvalue())
        self.assertNotIn('\x1b', private)

    def test_outlook_default_diagnostics_cover_unreadable_folder(self):
        class Folder:
            Name = 'Private Outlook folder'
            Folders = SimpleNamespace(Count=0)

            @property
            def Items(self):
                raise OSError('private@example.test')

        store = SimpleNamespace(FilePath=str(self.pst), GetRootFolder=Folder)
        namespace = SimpleNamespace(Stores=SimpleNamespace(Count=1, Item=lambda i: store))
        app = SimpleNamespace(GetNamespace=lambda name: namespace)
        client = SimpleNamespace(Dispatch=lambda name: app)
        modules = {'pythoncom': SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None),
                   'win32com': SimpleNamespace(client=client), 'win32com.client': client}
        console = StringIO()
        with patch.dict(sys.modules, modules), redirect_stdout(console), redirect_stderr(console):
            self.assertEqual(outlook.import_pst(self.db, self.pst, self.args), 2)
        self.assertEqual(self.db_path.with_suffix('.errors.txt').read_text(), 'folder.items error=OSError\n')
        for field in ('private@example.test', 'Private Outlook folder', str(self.pst)):
            self.assertNotIn(field, console.getvalue())

    def test_interrupted_import_can_resume(self):
        class Interrupted(FakeFolder):
            def get_sub_message(self, i):
                if i == 1:
                    raise KeyboardInterrupt
                return super().get_sub_message(i)

        messages = [FakeMessage(body='First'), FakeMessage(body='Second')]
        for folder, expected in [(Interrupted('Inbox', messages), 130), (FakeFolder('Inbox', messages), 0)]:
            module = SimpleNamespace(file=lambda: FakePFF(folder))
            self.assertEqual(direct.import_pst(self.db, self.pst, self.args, module), expected)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM messages').fetchone()[0], 2)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM source_items').fetchone()[0], 2)

    def test_changed_snapshot_rolls_back_current_batch_only(self):
        pst = self.pst

        class Changing(FakeFolder):
            def get_sub_message(self, i):
                if i == 2:
                    pst.write_bytes(b'a new snapshot during the third record')
                return super().get_sub_message(i)

        self.args.batch_size = 2
        module = SimpleNamespace(file=lambda: FakePFF(Changing('Inbox', [FakeMessage(body=str(i)) for i in range(4)])))
        with self.assertRaisesRegex(direct.ImportFailure, 'changed during import'):
            direct.import_pst(self.db, pst, self.args, module)
        self.assertEqual(self.db.execute('SELECT body FROM messages ORDER BY id').fetchall(), [('0',), ('1',)])
        self.assertFalse(self.db.in_transaction)

    def test_unrelated_existing_database_and_permissions_are_untouched(self):
        other = self.path / 'unrelated.db'
        with closing(sqlite3.connect(other)) as conn:
            conn.execute('CREATE TABLE tickets(id INTEGER)')
        before = other.stat().st_mode, other.read_bytes()
        with self.assertRaises(files.FileSafetyError):
            common.open_database(other)
        self.assertEqual((other.stat().st_mode, other.read_bytes()), before)
        self.assertFalse(Path(str(other) + '-wal').exists())

    def test_bad_reader_does_not_register_a_source_baseline(self):
        module = SimpleNamespace(file=lambda: FakePFF(None))
        with self.assertRaises(direct.ImportFailure):
            direct.import_pst(self.db, self.pst, self.args, module)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM import_sources').fetchone()[0], 0)

    def test_fts_display_escapes_control_sequences_without_mutating_evidence(self):
        item = FakeItem('a', '<id@example.test>', 'Inspection\x1b[2J', 'motor\ncase\u202e')
        message, _ = outlook.build_message(item, 'mail.pst', 'Inbox', False, False, 0)
        common.store_message(self.db, message)
        self.db.commit()
        console = StringIO()
        with redirect_stdout(console):
            common.search(self.db, 'motor', 1)
        self.assertNotIn('\x1b', console.getvalue())
        self.assertNotIn('\u202e', console.getvalue())
        self.assertIn('Inspection', console.getvalue())
        self.assertEqual(self.db.execute('SELECT subject, body FROM messages').fetchone(), (item.Subject, item.Body))

    def test_log_alias_of_source_is_refused_before_import(self):
        log = self.db_path.with_suffix('.direct.errors.txt')
        try:
            os.link(self.pst, log)
        except OSError:
            self.skipTest('Filesystem does not support hard links')
        source = self.pst.read_bytes()
        with self.assertRaises(SystemExit), redirect_stderr(StringIO()):
            direct.main(['--pst', str(self.pst), '--db', str(self.db_path)])
        self.assertEqual(self.pst.read_bytes(), source)

    def test_readonly_observes_committed_live_wal_and_closed_wal_stays_clean(self):
        self.db.execute('CREATE TABLE synthetic(value TEXT)')
        self.db.execute("INSERT INTO synthetic VALUES('from WAL')")
        self.db.commit()
        for writer_open in (True, False):
            if not writer_open:
                self.db.close()
            before = {p for p in self.path.iterdir()}
            with closing(files.open_readonly(self.db_path)) as reader:
                self.assertEqual(reader.execute('SELECT value FROM synthetic').fetchone(), ('from WAL',))
                with self.assertRaises(sqlite3.OperationalError):
                    reader.execute('DROP TABLE synthetic')
            self.assertEqual(set(self.path.iterdir()), before)

    @unittest.skipUnless(os.name == 'posix', 'POSIX mode bits')
    def test_private_database_log_and_sidecar_modes_under_permissive_umask(self):
        parent = self.path / 'existing'
        parent.mkdir(mode=0o755)
        before = parent.stat().st_mode
        archive = parent / 'new' / 'nested' / 'private.sqlite'
        previous = os.umask(0)
        try:
            with closing(common.open_database(archive)) as writer:
                writer.execute('CREATE TABLE synthetic(value TEXT)')
                with files.open_import_log(archive.with_suffix('.errors.txt')) as log:
                    log('item.read error=OSError')
                for p in (archive, archive.with_suffix('.errors.txt'), *files.sqlite_paths(archive)[1:3]):
                    self.assertEqual(p.stat().st_mode & 0o777, 0o600)
                for p in (archive.parent, archive.parent.parent):
                    self.assertEqual(p.stat().st_mode & 0o777, 0o700)
        finally:
            os.umask(previous)
        self.assertEqual(parent.stat().st_mode, before)

    def test_unreadable_folder_count_does_not_become_zero(self):
        class UnreadableFolder:
            @property
            def number_of_sub_messages(self):
                raise OSError('synthetic count failure')

        with self.assertRaisesRegex(OSError, 'synthetic count failure'):
            direct.folder_count(UnreadableFolder(), 'messages')


if __name__ == '__main__':
    unittest.main()
