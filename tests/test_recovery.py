"""Offline tests of metadata recovery. No real PST, mail, network, or Outlook."""
from __future__ import annotations
import hashlib
import contextlib
import io
from collections import Counter, defaultdict
from contextlib import closing
import json
import os
from pathlib import Path
import sqlite3
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
import pst_sqlite_recover as recover


class Entry:
    def __init__(self,pid,val,vt=0x001F):
        self.entry_type=pid
        self.value_type=vt
        self.val=val
    def get_data_as_string(self): return str(self.val)
    def get_data_as_integer(self): return int(self.val)


class RS:
    def __init__(self,props):self.entries=[Entry(*p) for p in props]
    @property
    def number_of_entries(self):return len(self.entries)
    def get_entry(self,i):return self.entries[i]


class Recs:
    def __init__(self,rows):self.rows=[RS(x) for x in rows]
    @property
    def number_of_recipients(self):return len(self.rows)
    def get_recipient(self,i):return self.rows[i]


class Msg:
    def __init__(self,props,recipients,body="",html=""):
        self.sets=[RS(props)]
        self.recipients=Recs(recipients)
        self.subject="Secret maintenance decision"
        self.plain_text_body=body
        self.html_body=html
        self.transport_headers=""
        self.sender_name="Confidential Technician"
        self.conversation_index=b"a"*22
    @property
    def number_of_record_sets(self):return len(self.sets)
    def get_record_set(self,i):return self.sets[i]


class Folder:
    def __init__(self,name,messages=(),folders=()):
        self.name=name
        self.messages=list(messages)
        self.folders=list(folders)
    @property
    def number_of_sub_messages(self):return len(self.messages)
    @property
    def number_of_sub_folders(self):return len(self.folders)
    def get_sub_message(self,i):return self.messages[i]
    def get_sub_folder(self,i):return self.folders[i]


class Reader:
    def __init__(self,root):self.root=root;self.closed=False
    def open(self,_):pass
    def get_root_folder(self):return self.root
    def close(self):self.closed=True


class Factory:
    def __init__(self,root):self.reader=Reader(root)
    def file(self):return self.reader


def insert_record(conn,pst,path,idx,mid,subject='',sender='',body='',to=''):
    conn.execute("""INSERT INTO messages(id,message_key,fingerprint,internet_message_id,conversation_key,
               sent_utc,received_utc,sender_name,sender_email,to_list,cc_list,subject,body,
               attachment_names_json,attachment_count,first_folder)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
             (mid,f'msg-{mid}',f'fp-{mid}','',f'conv-{mid}',
              '2026-11-02T10:00:00Z','2026-11-02T10:00:00Z','',sender,to,'',subject,body,'[]',0,path))
    key=hashlib.sha256((os.path.normcase(str(pst))+"\0"+path+"\0"+str(idx)).encode()).hexdigest()
    conn.execute("INSERT INTO source_items VALUES(?,?,?,?)",(key,mid,path,'kept'))
    conn.execute("INSERT INTO messages_fts(rowid,subject,body,sender_name,sender_email) VALUES(?,?,?,?,?)",(mid,subject,body,'',sender))


class RecoveryTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.directory=Path(self.temp.name)
        self.pst=self.directory/'archive.pst'
        self.pst.write_bytes(b'FAKE')
        self.old=self.directory/'old.sqlite'
        self.work=self.directory/'work.sqlite'
        self.log=self.directory/'audit.log'
        db=sqlite3.connect(self.old)
        db.executescript("""
            CREATE TABLE messages (
              id INTEGER PRIMARY KEY, message_key TEXT NOT NULL UNIQUE, fingerprint TEXT NOT NULL UNIQUE,
              internet_message_id TEXT,conversation_key TEXT,sent_utc TEXT,received_utc TEXT,
              sender_name TEXT,sender_email TEXT,to_list TEXT,cc_list TEXT,subject TEXT,body TEXT,
              attachment_names_json TEXT,attachment_count INTEGER,first_folder TEXT);
            CREATE TABLE source_items (
              source_key TEXT PRIMARY KEY, message_id INTEGER REFERENCES messages(id), folder TEXT,
              disposition TEXT);
            CREATE VIRTUAL TABLE messages_fts USING fts5(subject,body,sender_name,sender_email,
              content='messages',content_rowid='id');
        """)
        insert_record(db,self.pst,'PST/Sent',0,1)
        insert_record(db,self.pst,'PST/Archive',0,2)
        db.commit()
        db.close()
    def tearDown(self):self.temp.cleanup()

    def run_report(self,first,second):
        f=Factory(Folder('PST',folders=[Folder('Sent',[first]),Folder('Archive',[second])]))
        result=recover.run_recovery(self.pst,self.old,self.work,self.log,pypff_module=f)
        return result,f

    def test_recover_mapi_sender_and_typed_to_and_fts(self):
        first=Msg([(0x001A,'IPM.Note'),(0x5D01,'employee@example.org'),
                   (0x0037,'Secret maintenance decision'),(0x1035,'<message-id@example.org>')],
                  [[(0x0C15,1,3),(0x3001,'Engineering Lead'),(0x3003,'/o=EX/oU=site/cn=RECIP'),
                    (0x39FE,'engineering@example.org'),(0x3002,'EX')]],
                  body='',html='<html><body><p>Motor</p><p>Action approved</p></body></html>')
        second=Msg([(0x001A,'IPM.Appointment'),(0x5D02,'delegate@example.org')],[],body='')
        result,f=self.run_report(first,second)
        self.assertTrue(f.reader.closed)
        with closing(sqlite3.connect(self.old)) as old:
            self.assertEqual(old.execute('SELECT sender_email FROM messages WHERE id=1').fetchone()[0],'')
        with closing(sqlite3.connect(self.work)) as db:
            sender,to,body,cls,mid=db.execute("SELECT sender_email,to_list,body,message_category,internet_message_id FROM messages WHERE id=1").fetchone()
            self.assertEqual(sender,'employee@example.org')
            self.assertIn('engineering@example.org',to)
            self.assertIn('Action approved',body)
            self.assertEqual(cls,'Email / IPM.Note')
            self.assertEqual(mid,'message-id@example.org')
            self.assertEqual(db.execute("SELECT COUNT(*) FROM messages_fts WHERE messages_fts MATCH 'approved'").fetchone()[0],1)
            self.assertEqual(db.execute("SELECT sender_email FROM messages WHERE id=2").fetchone()[0],'')
            self.assertEqual(db.execute("SELECT represented_sender_smtp FROM recovery_metadata WHERE message_id=2").fetchone()[0],'delegate@example.org')
            self.assertEqual(db.execute("SELECT COUNT(*) FROM recovery_recipients").fetchone()[0],1)
        log=recover.safe_log(result,self.pst)
        for forbidden in ('employee@example.org','engineering@example.org','Secret maintenance decision','Confidential Technician','PST/Sent'):
            self.assertNotIn(forbidden,log)
        self.assertIn('filled_sender_smtp',log)

    def test_resolve_unique_exchange_alias_from_another_record(self):
        ex='/o=Corp/ou=Office/cn=PrivatePerson'
        # Sender only has EX address, and no mail header.
        first=Msg([(0x001A,'IPM.Note'),(0x0C1F,ex)],[],body='hello')
        # Another message's recipient table pairs the EX identifier with SMTP.
        second=Msg([(0x001A,'IPM.Note')],
                   [[(0x0C15,1,3),(0x3001,'Person'),(0x3002,'EX'),
                     (0x3003,ex),(0x39FE,'person@example.org')]],body='hello')
        result,f=self.run_report(first,second)
        with closing(sqlite3.connect(self.work)) as db:
            addr,method=db.execute("SELECT sender_email,sender_email_source FROM messages WHERE id=1").fetchone()
            self.assertEqual(addr,'person@example.org')
            self.assertEqual(method,'local_exchange_alias')
            self.assertEqual(db.execute("SELECT COUNT(*) FROM recovery_aliases").fetchone()[0],1)
        self.assertEqual(result['counts']['filled_sender_via_unique_alias'],1)

    def test_ambiguous_alias_not_used(self):
        ex='/o=Corp/ou=Office/cn=Alias'
        first=Msg([(0x001A,'IPM.Note'),(0x0C1F,ex)],[],body='hello')
        second=Msg([(0x001A,'IPM.Note')],[
            [(0x0C15,1,3),(0x3003,ex),(0x39FE,'first@example.org')],
            [(0x0C15,2,3),(0x3003,ex),(0x39FE,'second@example.org')],
        ])
        result,_=self.run_report(first,second)
        with closing(sqlite3.connect(self.work)) as db:
            self.assertEqual(db.execute('SELECT sender_email FROM messages WHERE id=1').fetchone()[0],'')
            self.assertEqual(db.execute('SELECT COUNT(*) FROM recovery_aliases').fetchone()[0],0)
        self.assertEqual(result['counts']['ambiguous_exchange_aliases'],1)

    def test_failed_write_rolls_back_fields_counters_and_alias_evidence(self):
        ex='/o=Corp/ou=Office/cn=Alias'
        first=Msg([(0x001A,'IPM.Note'),(0x0C1F,ex),(0x5D01,'person@example.org')],[],body='only failed row has text')
        second=Msg([(0x001A,'IPM.Note'),(0x0C1F,ex)],[])
        real_enrich=recover.enrich_one
        def fail_after_write(conn,mid,*args):
            real_enrich(conn,mid,*args)
            if mid==1:
                raise RuntimeError('private@example.org must never reach the log')
        old_hash=hashlib.sha256(self.old.read_bytes()).digest()
        with mock.patch.object(recover,'enrich_one',side_effect=fail_after_write):
            result,_=self.run_report(first,second)
        with closing(sqlite3.connect(self.work)) as db:
            self.assertEqual(db.execute('SELECT body FROM messages WHERE id=1').fetchone()[0],'')
            self.assertEqual(db.execute('SELECT COUNT(*) FROM recovery_metadata').fetchone()[0],1)
            self.assertEqual(db.execute('SELECT COUNT(*) FROM recovery_aliases').fetchone()[0],0)
            self.assertEqual(db.execute('SELECT sender_email FROM messages WHERE id=2').fetchone()[0],'')
        self.assertEqual(result['counts']['messages_recovered'],1)
        self.assertEqual(result['counts']['filled_body'],0)
        self.assertEqual(result['counts']['filled_sender_smtp'],0)
        self.assertEqual(result['counts']['filled_subject'],1)
        self.assertEqual(result['errors']['message.RuntimeError'],1)
        self.assertNotIn('private@example.org',recover.safe_log(result,self.pst))
        self.assertEqual(hashlib.sha256(self.old.read_bytes()).digest(),old_hash)

    def test_successful_savepoint_remains_in_outer_transaction(self):
        with closing(sqlite3.connect(self.old)) as db:
            recover.ensure_schema(db)
            recover.enrich_atomic(db,1,Msg([(0x001A,'IPM.Note')],[],body='pending'),
                                  Counter(),Counter(),defaultdict(Counter),4096,False)
            self.assertTrue(db.in_transaction)
            with closing(sqlite3.connect(self.old)) as observer:
                self.assertEqual(observer.execute('SELECT body FROM messages WHERE id=1').fetchone()[0],'')
            db.commit()
            with closing(sqlite3.connect(self.old)) as observer:
                self.assertEqual(observer.execute('SELECT body FROM messages WHERE id=1').fetchone()[0],'pending')

    def test_alias_evidence_counts_observations(self):
        ex='/o=Corp/ou=Office/cn=Alias'
        first=Msg([(0x001A,'IPM.Note'),(0x0C1F,ex),(0x5D01,'person@example.org')],[],body='hello')
        second=Msg([(0x001A,'IPM.Note')],[
            [(0x0C15,1,3),(0x3003,ex),(0x39FE,'Person@example.org')],
        ])
        self.run_report(first,second)
        with closing(sqlite3.connect(self.work)) as db:
            self.assertEqual(db.execute('SELECT evidence_count FROM recovery_aliases').fetchone()[0],2)

    def test_surrogate_repair_preserves_pairs_and_replaces_lone_units(self):
        first=Msg([(0x001A,'IPM.Note'),(0x0037,'Subject \ud83d\ude00'),
                   (0x0C1A,'Name \ud800')],
                  [[(0x0C15,1,3),(0x3001,'Recipient \ud800'),(0x3003,'person@example.org')]],
                  body='Before \ud83d\ude00 after \ud800 end')
        second=Msg([(0x001A,'IPM.Note')],[])
        result,_=self.run_report(first,second)
        with closing(sqlite3.connect(self.work)) as db:
            subject,body,name=db.execute('SELECT subject,body,sender_name FROM messages WHERE id=1').fetchone()
            self.assertEqual(subject,'Subject 😀')
            self.assertEqual(body,'Before 😀 after � end')
            self.assertEqual(name,'Name �')
            self.assertEqual(db.execute('SELECT display_name FROM recovery_recipients').fetchone()[0],'Recipient �')
        self.assertEqual(result['errors']['message.UnicodeEncodeError'],0)
        self.assertEqual(result['errors']['text.surrogate_fields_repaired'],4)

    def test_actual_rtf_decoder_and_unicode_storage(self):
        try:
            from striprtf.striprtf import rtf_to_text
        except ImportError:
            self.skipTest('Optional striprtf decoder is not installed')
        first=Msg([(0x001A,'IPM.Note')],[])
        first.rtf_body=r'{\rtf1\ansi Before \u-10179?\u-8704? after \u-10179? end}'
        second=Msg([(0x001A,'IPM.Note')],[])
        factory=Factory(Folder('PST',folders=[Folder('Sent',[first]),Folder('Archive',[second])]))
        result=recover.run_recovery(self.pst,self.old,self.work,self.log,use_rtf=True,pypff_module=factory)
        with closing(sqlite3.connect(self.work)) as db:
            body,source=db.execute('SELECT body,body_source FROM messages WHERE id=1').fetchone()
            self.assertEqual(body,'Before 😀 after � end')
            self.assertEqual(source,'rtf')
        self.assertEqual(result['counts']['body_source_rtf'],1)

    def test_mapi_html_fallback_reads_only_body_property(self):
        class BodyEntry(Entry):
            def get_data(self):
                return self.val
        class UnreadableHTML(Msg):
            @property
            def html_body(self):
                raise OSError('private subject hidden in exception')
            @html_body.setter
            def html_body(self,_):
                pass
        first=UnreadableHTML([(0x001A,'IPM.Note')],[])
        first.sets[0].entries.extend([
            Entry(0x3701,'attachment bytes must not be read',0x0102),
            Entry(0x1009,'compressed RTF must not be read here',0x0102),
            BodyEntry(0x1013,b'<p>Recovered HTML decision</p>',0x0102),
        ])
        second=Msg([(0x001A,'IPM.Note')],[])
        result,_=self.run_report(first,second)
        with closing(sqlite3.connect(self.work)) as db:
            self.assertEqual(db.execute('SELECT body_source FROM messages WHERE id=1').fetchone()[0],'mapi.html')
            self.assertEqual(db.execute('SELECT body FROM messages WHERE id=1').fetchone()[0],'Recovered HTML decision')
        self.assertEqual(result['counts']['body_source_mapi.html'],1)
        self.assertNotIn('private subject',recover.safe_log(result,self.pst))

    def test_header_sender_is_actual_from_is_represented(self):
        first=Msg([(0x001A,'IPM.Note')],[])
        first.transport_headers='Sender: assistant@example.org\nFrom: manager@example.org\n'
        second=Msg([(0x001A,'IPM.Note')],[])
        self.run_report(first,second)
        with closing(sqlite3.connect(self.work)) as db:
            self.assertEqual(db.execute('SELECT sender_email,sender_email_source FROM messages WHERE id=1').fetchone(),
                             ('assistant@example.org','header.sender'))
            self.assertEqual(db.execute('SELECT represented_sender_smtp FROM recovery_metadata WHERE message_id=1').fetchone()[0],
                             'manager@example.org')

    def test_existing_fields_identities_and_funes_map_are_preserved(self):
        with closing(sqlite3.connect(self.old)) as db:
            db.execute("UPDATE messages SET sender_email='original@example.org',sender_name='Original',to_list='to@example.org',cc_list='cc@example.org',subject='Original subject',body='Original body',internet_message_id='original-id' WHERE id=1")
            db.execute('CREATE TABLE funes_map(message_id INTEGER PRIMARY KEY,session_id TEXT,turn_uuid TEXT,seq INTEGER)')
            db.execute("INSERT INTO funes_map VALUES(1,'session-1','turn-1',0)")
            db.commit()
            columns='id,message_key,fingerprint,internet_message_id,conversation_key,sent_utc,received_utc,sender_name,sender_email,to_list,cc_list,subject,body,attachment_names_json,attachment_count,first_folder'
            expected=db.execute(f'SELECT {columns} FROM messages WHERE id=1').fetchone()
        first=Msg([(0x001A,'IPM.Note'),(0x5D01,'different@example.org'),(0x0037,'Different subject'),
                   (0x1035,'different-id')],
                  [[(0x0C15,1,3),(0x3003,'different-recipient@example.org')]],body='Different body')
        self.run_report(first,Msg([(0x001A,'IPM.Note')],[]))
        with closing(sqlite3.connect(self.work)) as db:
            self.assertEqual(db.execute(f'SELECT {columns} FROM messages WHERE id=1').fetchone(),expected)
            self.assertEqual(db.execute('SELECT * FROM funes_map').fetchone(),(1,'session-1','turn-1',0))
            self.assertEqual(db.execute('SELECT mapi_sender_smtp FROM recovery_metadata WHERE message_id=1').fetchone()[0],
                             'different@example.org')

    def test_changed_saved_snapshot_is_rejected_before_copy(self):
        signature=recover.source_signature(self.pst)
        with closing(sqlite3.connect(self.old)) as db:
            db.execute('CREATE TABLE import_sources(source_path TEXT PRIMARY KEY,size_bytes INTEGER,mtime_ns INTEGER)')
            db.execute('INSERT INTO import_sources VALUES(?,?,?)',(os.path.normcase(str(self.pst)),*signature[:2]))
            db.commit()
        self.pst.write_bytes(b'changed snapshot')
        with self.assertRaisesRegex(recover.RecoveryError,'saved import baseline'):
            self.run_report(Msg([],[]),Msg([],[]))
        self.assertFalse(self.work.exists())

    def test_changed_pst_during_scan_aborts_publication_and_writes_safe_log(self):
        pst=self.pst
        class ChangingFolder(Folder):
            def get_sub_message(self,i):
                pst.write_bytes(b'changed while reading')
                return super().get_sub_message(i)
        factory=Factory(Folder('PST',folders=[ChangingFolder('Sent',[Msg([(0x001A,'IPM.Note')],[])])]))
        with mock.patch.dict(sys.modules,{'pypff':factory}), contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            status=recover.main(['--pst',str(self.pst),'--db',str(self.old),
                                 '--out-db',str(self.work),'--log',str(self.log)])
        self.assertEqual(status,1)
        self.assertFalse(self.work.exists())
        self.assertEqual(list(self.directory.glob('.email-funes-recovery-*')),[])
        self.assertIn('PST changed during recovery',self.log.read_text())
        self.assertNotIn(str(self.pst),self.log.read_text())

    def test_fresh_input_required(self):
        self.run_report(Msg([(0x001A,'IPM.Note')],[]),Msg([(0x001A,'IPM.Note')],[]))
        with self.assertRaisesRegex(recover.RecoveryError,'already enriched'):
            recover.run_recovery(self.pst,self.work,self.directory/'second.sqlite',self.log)

    def test_publish_refuses_target_created_after_validation(self):
        self.work.write_bytes(b'completed output')
        target=self.directory/'target.sqlite'
        target.write_bytes(b'existing output')
        with self.assertRaises(FileExistsError):
            recover.publish_database(self.work,target)
        self.assertEqual(target.read_bytes(),b'existing output')
        self.assertEqual(self.work.read_bytes(),b'completed output')

    def test_log_cannot_overwrite_database_through_hardlink(self):
        try:
            os.link(self.old,self.log)
        except OSError:
            self.skipTest('Filesystem does not support hardlinks')
        original=self.old.read_bytes()
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as error:
            recover.main(['--pst',str(self.pst),'--db',str(self.old),
                          '--out-db',str(self.work),'--log',str(self.log)])
        self.assertEqual(error.exception.code,2)
        self.assertEqual(self.old.read_bytes(),original)

if __name__ == '__main__':
    unittest.main()
