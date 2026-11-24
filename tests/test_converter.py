"""Offline tests; Outlook is mocked, no PST or corporate data needed."""
import json
import sqlite3
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))

import pst_to_sqlite as p
import sqlite_to_funes as f


class FakeAccessor:
    def __init__(self, mid='', headers=''):
        self.mid = mid
        self.headers = headers
    def GetProperty(self, key):
        if '0x1035' in key:
            return self.mid
        if '0x007D' in key:
            return self.headers
        raise RuntimeError('missing MAPI property')


class FakeAttachments:
    def __init__(self, names):
        self.names = names
        self.Count = len(names)
    def Item(self, i):
        class Attachment:
            FileName = self.names[i-1]
        return Attachment()


class FakeItem:
    Class = 43
    def __init__(self, eid, mid, subject, body, folder='', headers='', conv='conversation-01', names=('one.pdf',)):
        self.EntryID = eid
        self.PropertyAccessor = FakeAccessor(mid, headers)
        self.Subject = subject
        self.Body = body
        self.SentOn = datetime(2026, 11, 1, 10, 0, 0, tzinfo=timezone.utc)
        self.ReceivedTime = datetime(2026, 11, 1, 10, 2, 0, tzinfo=timezone.utc)
        self.ConversationID = conv
        self.Attachments = FakeAttachments(names)
        self.SenderName = 'Technician'
        self.SenderEmailAddress = 'tech@example.com'
        self.To = 'engineer@example.com'
        self.CC = ''


class OfflineTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dbpath = Path(self.tmp.name) / 'mail.sqlite'
        self.db = p.open_database(self.dbpath)
    def tearDown(self):
        self.db.close()
        self.tmp.cleanup()

    def test_duplicate_by_message_id_different_folder(self):
        a = FakeItem('ENTRY-001', '<email-01@example.com>', 'Motor tests', 'Bearing replaced.')
        b = FakeItem('ENTRY-002', '<email-01@example.com>', 'Motor tests', 'Bearing replaced.')
        ma, reason = p.build_message(a, 'file.pst', 'Root/Inbox', False, False, 0)
        self.assertEqual(reason, '')
        mb, reason = p.build_message(b, 'file.pst', 'Root/Archive', False, False, 0)
        self.assertEqual(p.store_message(self.db, ma), 'kept')
        self.assertEqual(p.store_message(self.db, mb), 'duplicate')
        self.db.commit()
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM messages').fetchone()[0], 1)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM source_items').fetchone()[0], 2)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM messages_fts WHERE messages_fts MATCH ?', ('bearing',)).fetchone()[0], 1)

    def test_distinct_message_ids_are_preserved(self):
        a = FakeItem('a', '<a@example.com>', 'Motor tests', 'Bearing replaced.')
        b = FakeItem('b', '<b@example.com>', 'Motor tests', 'Bearing replaced.')
        self.assertEqual(p.store_message(self.db, p.build_message(a,'file.pst','Inbox',False,False,0)[0]), 'kept')
        self.assertEqual(p.store_message(self.db, p.build_message(b,'file.pst','Archive',False,False,0)[0]), 'kept')

    def test_spam_bulk_keep_and_folders(self):
        item = FakeItem('a','<a@example.com>','Special parts','details', headers='X-Spam-Flag: YES\r\n')
        self.assertEqual(p.build_message(item,'file.pst','Inbox',False,False,0)[1],'header-spam')
        self.assertEqual(p.build_message(item,'file.pst','Inbox',True,False,0)[1],'')
        bulk = FakeItem('b','<b@example.com>','Journal','newsletter',headers='List-Unsubscribe: <a@example.com>\r\n')
        self.assertEqual(p.build_message(bulk,'file.pst','Inbox',False,True,0)[1],'bulk-email')
        args=type('Args', (), {'exclude_junk': True,'exclude_deleted':True})()
        self.assertEqual(p.folder_exclusion('Courrier indésirable',args),'junk folder')
        self.assertEqual(p.folder_exclusion('Éléments supprimés',args),'deleted folder')

    def test_same_internet_id_with_distinct_content_is_preserved(self):
        a = FakeItem('a', '<a@example.com>', 'Decision', 'Version A')
        b = FakeItem('b', '<a@example.com>', 'Decision', 'Version B')
        m1, _ = p.build_message(a,'file.pst','Inbox',False,False,0)
        m2, _ = p.build_message(b,'file.pst','Archive',False,False,0)
        self.assertEqual(p.store_message(self.db,m1),'kept')
        self.assertEqual(p.store_message(self.db,m2),'kept')
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM messages').fetchone()[0],2)

    def test_attachment_names_only(self):
        item = FakeItem('a','<a@example.com>','Specs','Use the motor datasheet',names=('motor_data.pdf','photo.jpg'))
        msg, _ = p.build_message(item,'file.pst','Inbox',False,False,0)
        p.store_message(self.db,msg)
        self.assertEqual(json.loads(self.db.execute('SELECT attachment_names_json FROM messages').fetchone()[0]), ['motor_data.pdf','photo.jpg'])
        self.assertNotIn('PDF BINARY', str(self.db.execute('SELECT * FROM messages').fetchone()))

    def test_export_stable_funes_sequence(self):
        a = FakeItem('a','<a@example.com>','Motor tests','Check bearing.')
        a.SentOn = datetime(2026,11,1,tzinfo=timezone.utc)
        m,_ = p.build_message(a,'file.pst','Inbox',False,False,0)
        p.store_message(self.db,m)
        self.db.commit()
        f.ensure_map(self.db)
        self.assertEqual(f.map_new_messages(self.db, f.resolve_policy(self.db, include_all=True)),(1,0))
        out = Path(self.tmp.name)/'funes'
        self.assertEqual(f.export(self.db,out,per_file=2),1)
        records = [json.loads(line) for file in out.glob('*.jsonl') for line in file.read_text(encoding='utf8').splitlines()]
        self.assertEqual(records[0]['seq'],0)
        self.assertEqual(records[0]['harness'],'outlook-pst')
        self.assertIn('From:',records[0]['blocks'][0]['text'])
        # Insert earlier-timestamped email after first export, seq must append.
        b = FakeItem('b','<b@example.com>','Motor tests','Earlier mail.')
        b.SentOn = datetime(2026,10,1,tzinfo=timezone.utc)
        m,_=p.build_message(b,'file.pst','Archive',False,False,0)
        p.store_message(self.db,m)
        self.db.commit()
        self.assertEqual(f.map_new_messages(self.db),(1,0))
        self.assertEqual(f.map_new_messages(self.db),(0,0))
        out2 = Path(self.tmp.name)/'funes-second'
        self.assertEqual(f.export(self.db,out2,per_file=2),2)
        records = [json.loads(line) for file in out2.glob('*.jsonl') for line in file.read_text(encoding='utf8').splitlines()]
        self.assertEqual(sorted(r['seq'] for r in records),[0,1])
        self.assertTrue(records[0]['turn_uuid'].startswith('mid-'))


if __name__ == '__main__':
    unittest.main()
