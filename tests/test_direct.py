"""Offline unit tests: simulate pypff objects without accessing a real PST."""
from __future__ import annotations
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
import json
import tempfile
import unittest

import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))

import pst_to_sqlite_direct as d
import pst_to_sqlite as b
import sqlite_to_funes as f


class FakeAttachment:
    def __init__(self, name):
        self.long_filename = name
    def read_buffer(self, *args):
        raise AssertionError('Attachment payload must never be read')


class FakeMessage:
    def __init__(self, subject='Pump bearing', body='Bearing was replaced.',
                 headers='From: Tech <tech@example.org>\r\nTo: Engineer <eng@example.org>\r\nMessage-ID: <id-1@example.org>\r\n',
                 index=b'a'*22, filename='pump.pdf'):
        self.subject = subject
        self.plain_text_body = body
        self.html_body = ''
        self.transport_headers = headers
        self.sender_name = 'Tech'
        self.conversation_index = index
        self.client_submit_time = datetime(2026, 11, 2, 10, 0, 0)
        self.delivery_time = datetime(2026, 11, 2, 10, 0, 3)
        self.number_of_attachments = 1
        self.filename = filename
    def get_attachment(self, idx):
        assert idx == 0
        return FakeAttachment(self.filename)


class FakeFolder:
    def __init__(self, name, messages=(), folders=()):
        self.name = name
        self.messages = list(messages)
        self.subfolders = list(folders)
        self.number_of_sub_messages = len(self.messages)
        self.number_of_sub_folders = len(self.subfolders)
    def get_sub_message(self, idx):
        return self.messages[idx]
    def get_sub_folder(self, idx):
        return self.subfolders[idx]


class FakePFF:
    def __init__(self, root):
        self.root = root
        self.closed = False
        self.opened = None
    def open(self, path):
        self.opened = path
    def get_root_folder(self):
        return self.root
    def close(self):
        self.closed = True


class Tests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dbpath = Path(self.tmp.name) / 'mail.sqlite'
        self.db = b.open_database(self.dbpath)
        self.pst = Path(self.tmp.name) / 'archive.pst'
        self.pst.write_bytes(b'synthetic placeholder for mocked PST reader')
        self.args = SimpleNamespace(db=self.dbpath,limit=0,progress=1000,batch_size=2,
                                    exclude_junk=True,exclude_deleted=True,
                                    include_header_spam=False,skip_bulk=False,max_body_chars=0)

    def tearDown(self):
        self.db.close()
        self.tmp.cleanup()

    def test_message_attachment_names_and_dates(self):
        msg, reason = d.build_message(FakeMessage(), self.pst, 'Root/Inbox', 2, self.args)
        self.assertFalse(reason)
        self.assertEqual(msg.sent_utc, '2026-11-02T10:00:00Z')
        self.assertEqual(msg.sender_email, 'tech@example.org')
        self.assertIn('Engineer', msg.to_list)
        self.assertEqual(json.loads(msg.attachment_names_json), ['pump.pdf'])
        self.assertEqual(msg.attachment_count, 1)

    def test_exact_dedup_cross_folders(self):
        msg1, _ = d.build_message(FakeMessage(), self.pst, 'Root/Inbox', 0, self.args)
        msg2, _ = d.build_message(FakeMessage(), self.pst, 'Root/Sent', 1, self.args)
        self.assertEqual(b.store_message(self.db, msg1), 'kept')
        self.assertEqual(b.store_message(self.db, msg2), 'duplicate')
        self.assertEqual(self.db.execute('select count(*) from messages').fetchone()[0], 1)
        self.assertEqual(self.db.execute('select count(*) from source_items').fetchone()[0], 2)

    def test_spam_and_bulk_filter(self):
        spam = FakeMessage(headers='X-Spam-Flag: YES\r\nFrom: Tech <tech@example.org>\r\n')
        self.assertEqual(d.build_message(spam, self.pst, 'Root/Inbox', 0, self.args)[1], 'header-spam')
        bulk = FakeMessage(headers='List-Id: <newsletter.example>\r\n')
        self.args.skip_bulk = True
        self.assertEqual(d.build_message(bulk, self.pst, 'Root/Inbox', 0, self.args)[1], 'bulk-email')

    def test_html_fallback(self):
        msg = FakeMessage(body='')
        msg.html_body = '<html><style>.foo{}</style><body><p>Motor</p><p>Check bearing</p></body></html>'
        self.assertEqual(d.text_body(msg), 'Motor\n\nCheck bearing')

    def test_import_skips_junk_and_is_resumable(self):
        inbox = FakeFolder('Inbox', [FakeMessage('Run A'), FakeMessage('Run B', body='New content.')])
        junk = FakeFolder('Courrier indésirable', [FakeMessage('spam')])
        root = FakeFolder('Root', folders=[inbox, junk])
        pff = FakePFF(root)
        factory = SimpleNamespace(file=lambda: pff)
        self.assertEqual(d.import_pst(self.db, self.pst, self.args, factory), 0)
        self.assertTrue(pff.closed)
        self.assertEqual(self.db.execute('select count(*) from messages').fetchone()[0], 2)
        self.assertEqual(self.db.execute('select count(*) from source_items').fetchone()[0], 2)
        self.assertEqual(d.import_pst(self.db, self.pst, self.args, factory), 0)
        self.assertEqual(self.db.execute('select count(*) from messages').fetchone()[0], 2)

    def test_funes_export_compatibility(self):
        msg, _ = d.build_message(FakeMessage(), self.pst, 'Root/Inbox', 0, self.args)
        b.store_message(self.db, msg)
        self.db.commit()
        f.ensure_map(self.db)
        self.assertEqual(f.map_new_messages(self.db, f.resolve_policy(self.db, include_all=True)), (1,0))
        out = Path(self.tmp.name)/'funes'
        self.assertEqual(f.export(self.db,out),1)
        data = json.loads(next(out.glob('*.jsonl')).read_text(encoding='utf-8').strip())
        self.assertEqual(data['harness'], 'outlook-pst')
        self.assertIn('pump.pdf', data['blocks'][0]['text'])

    def test_shared_and_independent_mappings_for_two_same_named_archives(self):
        other = Path(self.tmp.name) / "elsewhere" / self.pst.name
        other.parent.mkdir()
        other.write_bytes(b"second synthetic archive")
        first = FakeMessage()
        second = FakeMessage(subject="Inspection follow-up", body="Alignment checked.")
        # Both archives contain the same conversation; the second repeats one
        # message. Folder and position alone cannot identify either occurrence.
        def load(db, path, messages):
            factory = SimpleNamespace(file=lambda: FakePFF(FakeFolder("Inbox", messages)))
            self.assertEqual(d.import_pst(db, path, self.args, factory), 0)

        def mapped(db):
            f.ensure_map(db)
            f.map_new_messages(db, f.resolve_policy(db, include_all=True))
            return list(f.turns(db))

        load(self.db, self.pst, [first])
        original = mapped(self.db)[0]
        load(self.db, other, [second, first])
        combined = mapped(self.db)
        self.assertEqual(combined[0], original)
        self.assertEqual([row["seq"] for row in combined], [0, 1])
        self.assertEqual(self.db.execute("SELECT COUNT(*) FROM source_items").fetchone()[0], 3)
        self.assertEqual(self.db.execute("SELECT COUNT(*) FROM import_sources").fetchone()[0], 2)
        # Independently mapped DBs have independent sequence allocators. Their
        # overlapping threads must use separate memories, or be imported into one DB.
        separate = b.open_database(Path(self.tmp.name) / "second.sqlite")
        try:
            load(separate, other, [second])
            independent = mapped(separate)[0]
        finally:
            separate.close()
        self.assertEqual(independent["session_id"], original["session_id"])
        self.assertEqual(independent["seq"], original["seq"])
        self.assertNotEqual(independent["turn_uuid"], original["turn_uuid"])
        self.assertEqual(independent["turn_uuid"], combined[1]["turn_uuid"])


if __name__ == '__main__':
    unittest.main()
