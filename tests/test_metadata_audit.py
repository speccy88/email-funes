"""Offline tests. No real PST, pypff binary, mailbox, or external data needed."""
import importlib.util
import contextlib
import io
from pathlib import Path
import hashlib
import sqlite3
import tempfile
import unittest
from contextlib import closing
import os

MODULE = Path(__file__).resolve().parents[1]/'scripts/pst_metadata_audit.py'
spec = importlib.util.spec_from_file_location('audit', MODULE)
audit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(audit)

class Entry:
    def __init__(self, pid, value, vt=0x001f):
        self.entry_type=pid
        self.value_type=vt
        self._value=value
    def get_data_as_string(self): return str(self._value)
    def get_data_as_integer(self): return int(self._value)

class RecordSet:
    def __init__(self, entries): self._entries=[Entry(*e) for e in entries]
    @property
    def number_of_entries(self): return len(self._entries)
    def get_entry(self, i): return self._entries[i]

class Recipients:
    def __init__(self, rows): self._rows=[RecordSet(x) for x in rows]
    @property
    def number_of_recipients(self): return len(self._rows)
    def get_recipient(self,i): return self._rows[i]

class Message:
    def __init__(self, headers='', sender='someone@example.org', rcpt='person@example.org', kind='IPM.Note'):
        self._rs=[RecordSet([
            (0x001a,kind),
            (0x5d01,sender),
            (0x0e04,rcpt),
            (0x1009,"RTF binary not read",0x0102),
        ])]
        self.recipients=Recipients([[
            (0x0c15,1,0x0003),
            (0x3003,rcpt),
            (0x39fe,rcpt),
            (0x3002,'SMTP')
        ]])
        self.transport_headers=headers
        self.sender_name='Display name private'
        self.conversation_index=b'a'*22
    @property
    def number_of_record_sets(self): return len(self._rs)
    def get_record_set(self,i): return self._rs[i]

class Folder:
    def __init__(self,name,items=(),folders=()):
        self.name=name
        self._items=list(items)
        self._folders=list(folders)
    @property
    def number_of_sub_messages(self): return len(self._items)
    def get_sub_message(self,i): return self._items[i]
    @property
    def number_of_sub_folders(self): return len(self._folders)
    def get_sub_folder(self,i): return self._folders[i]

class Reader:
    def __init__(self,root): self.root=root; self.was_closed=False
    def open(self,p): pass
    def get_root_folder(self): return self.root
    def close(self): self.was_closed=True

class Module:
    def __init__(self,root): self.reader=Reader(root)
    def file(self):return self.reader

class AuditTests(unittest.TestCase):
    def test_native_properties_recover_sender_and_recipients(self):
        errors=audit.Counter()
        out=audit.message_probe(Message(headers=''),errors)
        self.assertFalse(out['header_from'])
        self.assertTrue(out['mapi_sender_smtp'])
        self.assertTrue(out['sender_raw_smtp_like'])
        self.assertTrue(out['recipient_to'])
        self.assertTrue(out['recipient_smtp'])
        self.assertEqual(out['class'],'Email / IPM.Note')
        self.assertTrue(out['mapi_rtf_only'])
        self.assertEqual(errors.total(),0)

    def test_recipient_is_record_set_not_item(self):
        obj=RecordSet([(0x3003,'someone@example.org'),(0x0c15,2,0x0003)])
        fields,values,_=audit.extract_properties(obj,set(audit.RECIPIENT_PROPERTIES),audit.Counter(),'recipient')
        self.assertIn(0x3003,fields)
        self.assertEqual(values[0x0c15],'2')

    def test_log_excludes_private_metadata_and_paths(self):
        with tempfile.TemporaryDirectory() as td:
            pst=Path(td)/'private_name@company.test.pst'
            pst.write_bytes(b'synthetic-test')
            root=Folder('Private Secret Project',[Message('From: hidden@example.org\nSubject: Secret'), Message(kind='IPM.Contact')])
            stats=audit.scan(pst,None,0,0,Module(root))
            log=audit.report(stats,pst,False,0,'test',20)
            for private in ('hidden@example.org','someone@example.org','person@example.org','private_name','Secret Project','Secret'):
                self.assertNotIn(private,log)
            self.assertIn('Email / IPM.Note',log)
            self.assertIn('Contact / distribution list',log)
            self.assertEqual(stats.sampled,2)

    def test_spread_sampling(self):
        self.assertEqual(audit.sampled_indices(4,10),range(4))
        self.assertEqual(len(audit.sampled_indices(100,10)),10)
        self.assertEqual(audit.sampled_indices(3,0),range(3))

    def test_sqlite_comparison_read_only(self):
        with tempfile.TemporaryDirectory() as td:
            root_dir=Path(td)
            pst=root_dir/'archive.pst'
            pst.write_bytes(b'synthetic-test')
            dbfile=root_dir/'email.sqlite'
            folder_path='Private/archive/Inbox'
            key=hashlib.sha256((os.path.normcase(str(pst))+'\0'+folder_path+'\0'+'0').encode()).hexdigest()
            with closing(sqlite3.connect(dbfile)) as db:
                db.execute('CREATE TABLE messages (id INTEGER PRIMARY KEY, sender_email TEXT, to_list TEXT, subject TEXT, body TEXT)')
                db.execute('CREATE TABLE source_items (source_key TEXT, message_id INTEGER)')
                db.execute("INSERT INTO messages(id,sender_email,to_list,subject,body) VALUES(1,'','','','')")
                db.execute('INSERT INTO source_items VALUES(?,1)',(key,))
                db.commit()
            old=audit.open_old_db(dbfile,audit.Counter())
            try:
                root=Folder('Private',folders=[Folder('archive',folders=[Folder('Inbox',[Message()])])])
                stats=audit.scan(pst,old,0,0,Module(root))
                self.assertEqual(stats.db_compared,1)
                self.assertEqual(stats.db_missing_sender,1)
                self.assertEqual(stats.db_sender_potential,1)
                self.assertEqual(stats.db_to_potential,1)
            finally:
                old.close()

    def test_log_cannot_overwrite_source_through_hardlink(self):
        with tempfile.TemporaryDirectory() as td:
            pst=Path(td)/'archive.pst'
            pst.write_bytes(b'unchanged source')
            output=Path(td)/'report.log'
            try:
                os.link(pst,output)
            except OSError:
                self.skipTest('Filesystem does not support hardlinks')
            with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as error:
                audit.main(['--pst',str(pst),'--out',str(output)])
            self.assertEqual(error.exception.code,2)
            self.assertEqual(pst.read_bytes(),b'unchanged source')

if __name__=='__main__': unittest.main()
