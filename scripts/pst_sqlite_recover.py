#!/usr/bin/env python3
"""Enrich an existing email-funes SQLite archive from the original PST.

Never alters the original PST or SQLite file. Builds a fresh SQLite copy and
logs structural counts only; attachment binary contents are never accessed.

Requires: libpff-python (import pypff) and the accompanying
pst_metadata_audit.py from the previous, verified diagnostic.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from contextlib import closing
from email.message import Message
from email.parser import Parser
from email.utils import parseaddr
from html import unescape
from html.parser import HTMLParser
import base64
import binascii
import hashlib
import os
from pathlib import Path
import re
import sqlite3
import sys
import tempfile
import time
from typing import Any

import pst_metadata_audit as audit
from email_funes.adapters import PST_INSTALL_HELP
from email_funes_files import (FileSafetyError, create_private_database, diagnostic, new_private_text,
                              open_readonly, private_mkdir, require_distinct, source_signature, sqlite_paths)

VERSION = "1.1.0"
TARGET_MESSAGE_PROPS = {
    0x001A, 0x0037, 0x0070, 0x0C1A, 0x0C1F, 0x5D01, 0x5D02, 0x0065,
    0x0E04, 0x0E03, 0x0E02, 0x1035,
}
TARGET_RECIPIENT_PROPS = {0x0C15, 0x3001, 0x3002, 0x3003, 0x39FE}
SMTP_RE = re.compile(r"^[^\s<>@\x00-\x1f\x7f\ufffd]+@[^\s<>@\x00-\x1f\x7f\ufffd]+$")
SURROGATE_RE = re.compile(r"[\ud800-\udfff]")


class RecoveryError(RuntimeError):
    """A failure whose fixed message contains no mailbox content or path."""


def safe_text(value: Any, errors: Counter) -> str:
    """Combine UTF-16 surrogate pairs and replace isolated surrogates with U+FFFD.

    Some RTF parsers return UTF-16 code units in a Python string. SQLite's
    UTF-8 binding rejects those strings; valid supplementary characters must
    be combined rather than discarded. Never include the affected text in logs.
    """
    value = audit.text_value(value)
    if SURROGATE_RE.search(value):
        repaired = value.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "replace")
        errors["text.surrogate_fields_repaired"] += 1
        return repaired
    return value


def valid_smtp(value: str) -> bool:
    return bool(SMTP_RE.fullmatch(value.strip()))


def normalized_ex(value: str) -> str:
    value = value.strip()
    return value.casefold() if value.casefold().startswith(("/o=", "ex:")) else ""


def safe_header(value: str) -> Message:
    return Parser().parsestr(value, headersonly=True)


def mapi_message(msg: Any, errors: Counter):
    _present, values, _types = audit.extract_properties(msg, TARGET_MESSAGE_PROPS, errors, "mapi_message")
    return {key: safe_text(value, errors) for key, value in values.items()}


def read_recipients(msg: Any, errors: Counter, cap: int) -> tuple[list[dict], bool]:
    table = audit.try_attr(msg, "recipients", errors)
    if table is None:
        return [], False
    try:
        n = audit.get_number(table, "number_of_recipients", "get_number_of_recipients", errors)
    except Exception as exc:
        errors[f"recipient_count.{type(exc).__name__}"] += 1
        return [], False
    capped = n > cap
    if capped:
        errors["recipient_count_capped"] += 1
    rows = []
    for i in range(min(n, cap)):
        try:
            rec = table.get_recipient(i)
            _found, p, _types = audit.extract_properties(rec, TARGET_RECIPIENT_PROPS, errors, "recipient")
            typ = p.get(0x0C15, "")
            kind = {"1":"to", "2":"cc", "3":"bcc"}.get(typ, "unknown")
            name = safe_text(p.get(0x3001, ""), errors).strip()
            addr_type = safe_text(p.get(0x3002, ""), errors).strip()
            raw = safe_text(p.get(0x3003, ""), errors).strip()
            smtp_prop = safe_text(p.get(0x39FE, ""), errors).strip()
            smtp = smtp_prop if valid_smtp(smtp_prop) else (raw if valid_smtp(raw) else "")
            rows.append({"ordinal":i, "kind":kind, "display":name,
                         "raw":raw, "address_type":addr_type, "smtp":smtp})
        except Exception as exc:
            errors[f"recipient_row.{type(exc).__name__}"] += 1
    return rows, capped


def display_recipient(r: dict) -> str:
    name = r["display"].replace("\r", " ").replace("\n", " ").strip()
    smtp = r["smtp"]
    if smtp and name and name.casefold() != smtp.casefold():
        return f"{name} <{smtp}>"
    return smtp or name


def join_role(rows: list[dict], kind: str) -> str:
    ordered = []
    used = set()
    for row in rows:
        if row["kind"] != kind:
            continue
        value = display_recipient(row)
        if value and value.casefold() not in used:
            ordered.append(value)
            used.add(value.casefold())
    return "; ".join(ordered)


class TextHTML(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.hide = 0
    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self.hide += 1
        elif tag in ("br", "p", "div", "li", "tr", "h1", "h2", "h3"):
            self.parts.append("\n")
    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self.hide = max(0, self.hide-1)
        elif tag in ("p", "div", "li", "tr"):
            self.parts.append("\n")
    def handle_data(self, data):
        if not self.hide:
            self.parts.append(data)


def clean_text(value: Any, errors: Counter) -> str:
    s = safe_text(value, errors)
    s = s.replace("\x00", "").replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"[ \t]+\n", "\n", s)
    return re.sub(r"\n{5,}", "\n\n\n", s).strip()


def html_text(raw: Any, errors: Counter) -> str:
    parser = TextHTML()
    parser.feed(audit.text_value(raw))
    parser.close()
    return clean_text(unescape("".join(parser.parts)), errors)


def mapi_body_fallback(msg: Any, errors: Counter) -> tuple[str, str]:
    """Read only PR_BODY / PR_HTML when the high-level body accessors fail.

    Binary PR_HTML is a message body, not an attachment. No other binary MAPI
    property is read, including compressed RTF and attachment payloads.
    """
    try:
        nsets = audit.get_number(msg, "number_of_record_sets", "get_number_of_record_sets", errors)
        for si in range(min(nsets, audit.MAX_RECORD_SETS)):
            rs = msg.get_record_set(si)
            total = audit.get_number(rs, "number_of_entries", "get_number_of_entries", errors)
            for ei in range(min(total, audit.MAX_ENTRIES_PER_SET)):
                entry = rs.get_entry(ei)
                pid = audit.property_id(audit.try_attr(entry, "entry_type", errors))
                if pid not in (0x1000, 0x1013):
                    continue
                try:
                    value_type = int(audit.try_attr(entry, "value_type", errors) or 0)
                    if value_type in audit.STRING_VALUE_TYPES:
                        raw = entry.get_data_as_string()
                    elif pid == 0x1013 and value_type == 0x0102:
                        raw = entry.get_data()
                    else:
                        continue
                    text = clean_text(raw, errors) if pid == 0x1000 else html_text(raw, errors)
                    if text:
                        return text, "mapi.plain" if pid == 0x1000 else "mapi.html"
                except Exception as exc:
                    errors[f"body.mapi.{hex(pid)}.{type(exc).__name__}"] += 1
    except Exception as exc:
        errors[f"body.mapi_scan.{type(exc).__name__}"] += 1
    return "", ""


def recover_body(msg: Any, errors: Counter, use_rtf: bool) -> tuple[str, str]:
    """Return only successfully decoded body text and its source."""
    for field in ("plain_text_body", "html_body"):
        try:
            raw = getattr(msg, field, None)
            if not raw:
                continue
            if field == "plain_text_body":
                text = clean_text(raw, errors)
            else:
                text = html_text(raw, errors)
            if text:
                return text, "plain" if field == "plain_text_body" else "html"
        except Exception as exc:
            errors[f"body.{field}.{type(exc).__name__}"] += 1
    text, source = mapi_body_fallback(msg, errors)
    if text:
        return text, source
    if use_rtf:
        try:
            from striprtf.striprtf import rtf_to_text
        except ImportError:
            errors["rtf.decoder_unavailable"] += 1
            return "", ""
        try:
            raw = getattr(msg, "rtf_body", None)
            if raw:
                payload = audit.text_value(raw)
                if payload.lstrip().startswith("{\\rtf"):
                    text = clean_text(rtf_to_text(payload), errors)
                    if text:
                        return text, "rtf"
                else:
                    errors["rtf.unrecognized_payload"] += 1
        except Exception as exc:
            errors[f"rtf.decode.{type(exc).__name__}"] += 1
    return "", ""


def best_sender(props: dict, h, errors: Counter) -> tuple[str, str, str, str]:
    """Actual sender only: represented sender stays in a separate field."""
    mapi_sender = props.get(0x0C1F, "").strip()
    mapi_smtp = props.get(0x5D01, "").strip()
    represented = props.get(0x5D02, "").strip() or props.get(0x0065, "").strip()
    _name, header_sender = parseaddr(h.get("Sender", ""))
    _name, header_from = parseaddr(h.get("From", ""))
    if not valid_smtp(represented) and valid_smtp(header_sender) and valid_smtp(header_from):
        if header_sender.casefold() != header_from.casefold():
            represented = header_from
    candidates = ((mapi_smtp, "mapi.sender_smtp"),
                  (mapi_sender, "mapi.sender_email"),
                  (header_sender, "header.sender"),
                  (header_from, "header.from"))
    for val, label in candidates:
        if valid_smtp(val):
            return val, label, mapi_sender, represented if valid_smtp(represented) else ""
    return "", "", mapi_sender or mapi_smtp, represented if valid_smtp(represented) else ""


def kind_of_item(props: dict, msg: Any, errors: Counter) -> tuple[str, str]:
    raw = props.get(0x001A, "") or safe_text(audit.try_attr(msg,"message_class", errors), errors)
    return raw, audit.item_class_bucket(raw)


def thread_candidate(msg: Any, headers, errors: Counter) -> str:
    idx = audit.try_attr(msg, "conversation_index", errors)
    if not isinstance(idx, (bytes, bytearray)) or len(idx) < 22 or (len(idx)-22)%5:
        try:
            idx = base64.b64decode("".join(str(headers.get("Thread-Index", "")).split()), validate=True)
        except (ValueError, binascii.Error):
            idx = b""
    if isinstance(idx, (bytes,bytearray)) and len(idx)>=22 and (len(idx)-22)%5==0:
        # Match the direct importer's established thread-key convention.
        return "outlook-"+hashlib.sha256(bytes(idx[:22]).hex().encode("ascii")).hexdigest()[:32]
    return ""  # Existing stored thread ID is deliberately left untouched.


def ensure_schema(conn: sqlite3.Connection):
    required = {"id", "message_key", "sender_email", "to_list", "cc_list", "subject", "body", "internet_message_id"}
    present = {r[1] for r in conn.execute("PRAGMA table_info(messages)")}
    if not required.issubset(present):
        raise ValueError("This is not a compatible email-funes SQLite database")
    conn.execute("SELECT source_key,message_id FROM source_items LIMIT 1")
    conn.execute("SELECT rowid FROM messages_fts LIMIT 1")
    for key in ("message_class", "message_category", "sender_raw_address", "sender_email_source", "recipient_to_source", "body_source"):
        if key not in present:
            conn.execute(f"ALTER TABLE messages ADD COLUMN {key} TEXT NOT NULL DEFAULT ''")
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS recovery_metadata (
            message_id INTEGER PRIMARY KEY REFERENCES messages(id),
            mapi_sender_raw TEXT NOT NULL DEFAULT '',
            mapi_sender_smtp TEXT NOT NULL DEFAULT '',
            represented_sender_smtp TEXT NOT NULL DEFAULT '',
            display_to TEXT NOT NULL DEFAULT '',
            display_cc TEXT NOT NULL DEFAULT '',
            display_bcc TEXT NOT NULL DEFAULT '',
            recovered_thread_key TEXT NOT NULL DEFAULT '',
            recovered_internet_id TEXT NOT NULL DEFAULT '',
            recipient_count INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS recovery_recipients (
            message_id INTEGER NOT NULL REFERENCES messages(id),
            ordinal INTEGER NOT NULL,
            recipient_type TEXT NOT NULL,
            display_name TEXT NOT NULL,
            address_type TEXT NOT NULL,
            address_raw TEXT NOT NULL,
            smtp_address TEXT NOT NULL,
            PRIMARY KEY (message_id, ordinal)
        );
        CREATE TABLE IF NOT EXISTS recovery_aliases (
            exchange_address TEXT PRIMARY KEY,
            smtp_address TEXT NOT NULL,
            evidence_count INTEGER NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_recovery_recipients_message ON recovery_recipients(message_id);
        CREATE INDEX IF NOT EXISTS idx_messages_recovery_category ON messages(message_category);
    """)
    conn.commit()


def stats_snapshot(conn):
    return {
        "rows": conn.execute("SELECT COUNT(*) FROM messages").fetchone()[0],
        "sender": conn.execute("SELECT COUNT(*) FROM messages WHERE TRIM(COALESCE(sender_email,''))='' ").fetchone()[0],
        "to": conn.execute("SELECT COUNT(*) FROM messages WHERE TRIM(COALESCE(to_list,''))='' ").fetchone()[0],
        "cc": conn.execute("SELECT COUNT(*) FROM messages WHERE TRIM(COALESCE(cc_list,''))='' ").fetchone()[0],
        "body": conn.execute("SELECT COUNT(*) FROM messages WHERE TRIM(COALESCE(body,''))='' ").fetchone()[0],
        "subject": conn.execute("SELECT COUNT(*) FROM messages WHERE TRIM(COALESCE(subject,''))='' ").fetchone()[0],
        "message_id": conn.execute("SELECT COUNT(*) FROM messages WHERE TRIM(COALESCE(internet_message_id,''))='' ").fetchone()[0],
    }


def add_pair(aliases: dict[str, Counter], raw: str, smtp: str):
    key = normalized_ex(raw)
    if key and valid_smtp(smtp):
        aliases[key][smtp.casefold()] += 1


def enrich_one(conn, message_id, msg, errors: Counter, counts: Counter,
               aliases: dict[str, Counter], max_recipients: int, use_rtf: bool):
    """Update one unique stored email in the *copy*, preserving stable IDs."""
    p = mapi_message(msg, errors)
    klass, category = kind_of_item(p, msg, errors)
    raw_headers = safe_text(audit.try_attr(msg, "transport_headers", errors), errors)
    h = safe_header(raw_headers)
    sender_smtp, sender_source, raw_sender, represented = best_sender(p, h, errors)
    rows, capped = read_recipients(msg, errors, max_recipients)
    if capped:
        counts["recipient_table_capped_items"] += 1
    for row in rows:
        add_pair(aliases, row["raw"], row["smtp"])
    add_pair(aliases, p.get(0x0C1F, ""), p.get(0x5D01, ""))
    add_pair(aliases, p.get(0x0065, ""), p.get(0x5D02, ""))
    db_row = conn.execute("SELECT sender_email,to_list,cc_list,subject,body,sender_name,internet_message_id FROM messages WHERE id=?",(message_id,)).fetchone()
    if db_row is None:
        raise ValueError("Source occurrence references missing message")
    old_sender, old_to, old_cc, old_subject, old_body, old_sender_name, old_mid = db_row
    old_sender, old_to, old_cc, old_subject, old_body = [x or "" for x in (old_sender,old_to,old_cc,old_subject,old_body)]
    new_sender = sender_smtp if not old_sender.strip() else ""
    to_candidates = join_role(rows,"to")
    cc_candidates = join_role(rows,"cc")
    # Display-To is often names, not an email address. Mark the provenance.
    display_to = p.get(0x0E04, "").strip()
    display_cc = p.get(0x0E03, "").strip()
    display_bcc = p.get(0x0E02, "").strip()
    to_value = (to_candidates or display_to) if not old_to.strip() else ""
    cc_value = (cc_candidates or display_cc) if not old_cc.strip() else ""
    # Native message class/subject may be available even without transport headers.
    new_subject = (p.get(0x0037, "") or safe_text(audit.try_attr(msg, "subject", errors), errors)).strip() if not old_subject.strip() else ""
    new_sender_name = (p.get(0x0C1A, "") or safe_text(audit.try_attr(msg, "sender_name", errors), errors)).strip() if not str(old_sender_name or "").strip() else ""
    new_mid = (p.get(0x1035, "") or h.get("Message-ID", "")).strip(" <>\r\n") if not str(old_mid or "").strip() else ""
    # Email-only body restoration. Avoid converting appointments/contacts to email text.
    recovered_body, body_source = ("", "")
    if not old_body.strip() and category in ("Email / IPM.Note", "Meeting request"):
        recovered_body, body_source = recover_body(msg, errors, use_rtf)
    changes = []
    vals = []
    for field, value in (("message_class",klass),("message_category",category),
                         ("sender_raw_address",raw_sender),
                         ("sender_email_source", sender_source if new_sender else ""),
                         ("recipient_to_source", ("recipient.table" if to_candidates else "mapi.display_to") if to_value else ""),
                         ("body_source",body_source)):
        changes.append(field+"=?")
        vals.append(value)
    for field, value, key in (("sender_email",new_sender,"filled_sender_smtp"),
                              ("to_list",to_value,"filled_to"),("cc_list",cc_value,"filled_cc"),
                              ("subject",new_subject,"filled_subject"),
                              ("sender_name",new_sender_name,"filled_sender_name"),
                              ("internet_message_id",new_mid,"filled_message_id"),
                              ("body",recovered_body,"filled_body")):
        if value:
            changes.append(field+"=?")
            vals.append(value)
            counts[key] += 1
            if field=="body": counts["body_source_"+body_source]+=1
    vals.append(message_id)
    conn.execute("UPDATE messages SET " + ", ".join(changes) + " WHERE id=?", vals)
    conn.execute("""INSERT INTO recovery_metadata(message_id,mapi_sender_raw,mapi_sender_smtp,represented_sender_smtp,display_to,display_cc,display_bcc,recovered_thread_key,recovered_internet_id,recipient_count) VALUES(?,?,?,?,?,?,?,?,?,?)""",
                 (message_id,raw_sender,sender_smtp,represented,display_to,display_cc,display_bcc,thread_candidate(msg,h,errors),new_mid,len(rows)))
    for row in rows:
        conn.execute("""INSERT INTO recovery_recipients(message_id,ordinal,recipient_type,display_name,address_type,address_raw,smtp_address) VALUES(?,?,?,?,?,?,?)""",
                     (message_id,row["ordinal"],row["kind"],row["display"],row["address_type"],row["raw"],row["smtp"]))
    counts["class_"+category]+=1
    counts["recipients_saved"]+=len(rows)
    if not valid_smtp(sender_smtp) and normalized_ex(raw_sender):
        counts["sender_exchange_unresolved_initial"]+=1


def enrich_atomic(conn, message_id, msg, errors: Counter, counts: Counter,
                  aliases: dict[str, Counter], max_recipients: int, use_rtf: bool):
    """Keep field counters and alias evidence consistent with successful writes."""
    local_counts = Counter()
    local_aliases = defaultdict(Counter)
    # A SAVEPOINT without an outer transaction commits when released. Keep a
    # real batch transaction so the caller controls disk commits.
    if not conn.in_transaction:
        conn.execute("BEGIN")
    conn.execute("SAVEPOINT one_recovery")
    try:
        enrich_one(conn, message_id, msg, errors, local_counts, local_aliases,
                   max_recipients, use_rtf)
        conn.execute("RELEASE one_recovery")
    except BaseException:
        conn.execute("ROLLBACK TO one_recovery")
        conn.execute("RELEASE one_recovery")
        raise
    counts.update(local_counts)
    for raw, observations in local_aliases.items():
        aliases[raw].update(observations)
    counts["messages_recovered"] += 1


def resolve_aliases(conn, aliases: dict[str, Counter], counts: Counter):
    for raw, possibles in aliases.items():
        if len(possibles)==1:
            smtp=next(iter(possibles))
            conn.execute("INSERT INTO recovery_aliases(exchange_address,smtp_address,evidence_count) VALUES(?,?,?)",(raw,smtp,possibles[smtp]))
        else:
            counts["ambiguous_exchange_aliases"]+=1
    for mid, raw in conn.execute("""SELECT m.id, m.sender_raw_address FROM messages m
                               JOIN recovery_metadata x ON x.message_id=m.id
                               WHERE TRIM(COALESCE(m.sender_email,''))='' AND TRIM(m.sender_raw_address)<>''"""):
        key=normalized_ex(raw)
        if not key or len(aliases.get(key,()))!=1:
            continue
        smtp=next(iter(aliases[key]))
        conn.execute("UPDATE messages SET sender_email=?,sender_email_source='local_exchange_alias' WHERE id=?",(smtp,mid))
        counts["filled_sender_via_unique_alias"]+=1


def check_source_baseline(db, pst: Path, signature: tuple, counts: Counter):
    """Verify the importer's saved snapshot when the source database has one."""
    has_table = db.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='import_sources'"
    ).fetchone()
    if has_table:
        previous = db.execute(
            "SELECT size_bytes,mtime_ns FROM import_sources WHERE source_path=?",
            (os.path.normcase(str(pst)),),
        ).fetchone()
        if previous is not None:
            if tuple(previous) != signature[:2]:
                raise RecoveryError("The PST size or modification time differs from the saved import baseline. Use the unchanged original PST.")
            counts["source_baseline_verified"] = 1
            return
        if db.execute("SELECT 1 FROM import_sources LIMIT 1").fetchone():
            raise RecoveryError("The PST path does not match a saved import source. Use the original PST path on the original operating system.")
    counts["source_baseline_unavailable"] = 1


def require_unchanged_source(pst: Path, signature: tuple):
    if source_signature(pst) != signature:
        raise RecoveryError("The PST changed during recovery. Close applications using it and retry with the unchanged original.")


def run_recovery(pst:Path, original:Path, working:Path, log:Path,
                 limit_matched:int=0, max_recipients:int=4096, progress:int=4000,
                 use_rtf:bool=False, pypff_module=None) -> dict:
    errors=Counter()
    counts=Counter()
    aliases: dict[str,Counter]=defaultdict(Counter)
    started=time.monotonic()
    pst, original, working = pst.resolve(), original.resolve(), working.absolute()
    require_distinct(pst, *sqlite_paths(original), *sqlite_paths(working), log)
    if working in (pst, original) or (working.exists() and working.stat().st_size):
        raise RecoveryError("Recovery needs a separate empty temporary database.")
    signature=source_signature(pst)
    # A consistent SQLite backup copies even a source currently using WAL.
    with closing(open_readonly(original)) as src:
        if src.execute("SELECT 1 FROM sqlite_master WHERE name IN ('recovery_metadata','recovery_recipients','recovery_aliases') LIMIT 1").fetchone():
            raise RecoveryError("The input database was already enriched. Start from the original pre-recovery database and use a new output name.")
        check_source_baseline(src,pst,signature,counts)
        create_private_database(working)
        dst=sqlite3.connect(working)
        try:
            src.backup(dst)
        finally:
            dst.close()
    db=sqlite3.connect(working,timeout=120)
    db.execute("PRAGMA foreign_keys=ON")
    db.execute("PRAGMA busy_timeout=120000")
    db.execute("PRAGMA journal_mode=DELETE")
    file=None
    completed=False
    try:
        ensure_schema(db)
        before=stats_snapshot(db)
        if pypff_module is None:
            import pypff as pypff_module
        file=pypff_module.file()
        file.open(str(pst))
        root=file.get_root_folder()
        if root is None:
            raise RuntimeError("PST has no readable root folder")
        root_name=audit.text_value(audit.try_attr(root,"name",errors)) or "PST"
        stack=[(root,root_name)]
        while stack:
            folder,path=stack.pop()
            counts["folders_seen"]+=1
            try:
                nchild=audit.get_number(folder,"number_of_sub_folders","get_number_of_sub_folders",errors)
                for i in range(nchild-1,-1,-1):
                    try:
                        child=folder.get_sub_folder(i)
                        stack.append((child,path+"/"+(audit.text_value(audit.try_attr(child,"name",errors)) or str(i))))
                    except Exception as exc:
                        errors["child."+type(exc).__name__]+=1
            except Exception as exc:
                errors["folder_children."+type(exc).__name__]+=1
            try:
                count=audit.get_number(folder,"number_of_sub_messages","get_number_of_sub_messages",errors)
            except Exception as exc:
                errors["folder_messages."+type(exc).__name__]+=1
                continue
            for i in range(count):
                counts["source_slots"]+=1
                key=hashlib.sha256((os.path.normcase(str(pst))+"\0"+path+"\0"+str(i)).encode("utf-8","replace")).hexdigest()
                row=db.execute("SELECT message_id FROM source_items WHERE source_key=?",(key,)).fetchone()
                if not row:
                    counts["unmatched_slots"]+=1
                    continue
                counts["matched_slots"]+=1
                mid=int(row[0])
                if db.execute("SELECT 1 FROM recovery_metadata WHERE message_id=?",(mid,)).fetchone():
                    counts["duplicate_sources_already_recovered"]+=1
                    continue
                try:
                    msg=folder.get_sub_message(i)
                    enrich_atomic(db,mid,msg,errors,counts,aliases,max_recipients,use_rtf)
                except Exception as exc:
                    errors["message."+type(exc).__name__]+=1
                if counts["messages_recovered"] and counts["messages_recovered"]%250==0:
                    db.commit()
                if progress and counts["source_slots"]%progress==0:
                    print(f"Inspected {counts['source_slots']:,} PST slots; enriched {counts['messages_recovered']:,} unique SQLite rows",flush=True)
                if limit_matched and counts["messages_recovered"]>=limit_matched:
                    counts["partial_limit_reached"]+=1
                    break
            db.commit()
            if limit_matched and counts["partial_limit_reached"]:
                break
        resolve_aliases(db,aliases,counts)
        db.commit()
        if before["rows"] and not counts["matched_slots"]:
            raise RecoveryError("No PST positions matched the database. Use the same unchanged PST, original path, and operating system as the first import.")
        require_unchanged_source(pst,signature)
        # Recreate FTS after all field replacements, avoiding stale FTS5 tokens.
        db.execute("INSERT INTO messages_fts(messages_fts) VALUES('rebuild')")
        db.commit()
        after=stats_snapshot(db)
        quick=db.execute("PRAGMA quick_check").fetchone()[0]
        fk=len(db.execute("PRAGMA foreign_key_check").fetchall())
        counts["known_message_rows"]=before["rows"]
        counts["fts_entries"]=db.execute("SELECT COUNT(*) FROM messages_fts").fetchone()[0]
        if quick!="ok" or fk:
            raise RuntimeError("Recovered SQLite failed integrity verification")
        # Close the reader before the final signature check. No PST write is
        # requested, and a changing source must never be published as valid.
        try:
            file.close()
        except Exception:
            errors["pst.close"]+=1
        file=None
        require_unchanged_source(pst,signature)
        completed=True
        outcome={"before":before,"after":after,"counts":counts,"errors":errors,
                 "elapsed_seconds":round(time.monotonic()-started,1),
                 "sqlite_integrity":quick,"foreign_key_failures":fk,"partial":bool(counts["partial_limit_reached"])}
        return outcome
    finally:
        if file is not None:
            try:file.close()
            except Exception: errors["pst.close"]+=1
        if not completed:
            db.rollback()
        db.close()


def safe_log(outcome:dict,pst:Path) -> str:
    counts=outcome["counts"]
    before,after=outcome["before"],outcome["after"]
    rows=["PST -> SQLITE METADATA RECOVERY REPORT","="*54,
          f"Script version: {VERSION}","Original PST filename/path: [REDACTED]",
          f"PST size (GiB): {pst.stat().st_size/2**30:.2f}",
          "Original SQLite filename/path: [REDACTED]",
          "Output SQLite filename/path: [REDACTED]",
          "Contains email bodies, addresses, subjects or folder paths: NO",
          f"Elapsed seconds: {outcome['elapsed_seconds']}",
          f"Import scope: {'PARTIAL TEST' if outcome['partial'] else 'FULL SCAN'}",
          "", "[SOURCE MATCHING]",
          f"PST folders inspected: {counts['folders_seen']}",
          f"PST message positions examined: {counts['source_slots']}",
          f"Source positions matched to old SQLite: {counts['matched_slots']}",
          f"Source positions without SQLite match: {counts['unmatched_slots']}",
          f"Unique messages enriched: {counts['messages_recovered']}",
          f"Duplicate source positions: {counts['duplicate_sources_already_recovered']}",
          f"Recipients stored in normalized table: {counts['recipients_saved']}",
          "", "[MISSING FIELDS BEFORE AND AFTER]",
          "Field                   before        after   filled",
         ]
    for name in ("sender","to","cc","subject","body","message_id"):
        b=before[name];a=after[name]
        rows.append(f"{name:22s} {b:9,d} {a:12,d} {b-a:8,d}")
    rows += ["", "[RECOVERY SOURCES]"]
    for name in ("filled_sender_smtp","filled_sender_via_unique_alias","filled_to","filled_cc","filled_subject","filled_message_id","filled_sender_name","filled_body","body_source_plain","body_source_html","body_source_mapi.plain","body_source_mapi.html","body_source_rtf","recipient_table_capped_items","sender_exchange_unresolved_initial","ambiguous_exchange_aliases"):
        rows.append(f"{name:38s} {counts[name]:,}")
    rows += ["", "[MESSAGE CLASSES IN ENRICHED SOURCE ROWS]"]
    for k,n in sorted(counts.items(),key=lambda x:-x[1]):
        if k.startswith("class_"):
            rows.append(f"{k[6:36]:36s} {n:,}")
    rows += ["", "[VALIDATION]",
             f"SQLite quick_check: {outcome['sqlite_integrity']}",
             f"Foreign key violations: {outcome['foreign_key_failures']}",
             f"FTS records: {counts['fts_entries']}",
             f"Original import snapshot baseline: {'verified' if counts['source_baseline_verified'] else 'unavailable in legacy database'}",
             "PST size, modification time and file identity: unchanged during this pass",
             "Message IDs, fingerprints, and Funes map (if any): unchanged",
             "Original SQLite/PST: opened read-only",
             "Recovery counters and alias evidence: successful message writes only",
             "", "[SANITIZED ERROR COUNTS]"]
    if outcome["errors"]:
        for label,count in outcome["errors"].most_common(30):
            rows.append(f"{label}: {count:,}")
    else:
        rows.append("None")
    rows += ["", "[LIMITATIONS]",
             "This pass matches source positions from the previous importer; use the original path and operating system.",
             "A legacy database without an import snapshot cannot prove the PST was unchanged before this pass.",
             "Display-To can contain names rather than email addresses. Recipient rows preserve source/provenance.",
             "Only unique locally observed Exchange->SMTP pairs are used for alias resolution.",
             "Alias evidence_count records observed property pairs, not unique people or messages.",
             "Represented sender addresses are recorded separately, never substituted for actual sender.",
             "Isolated Unicode surrogates are replaced with U+FFFD; paired code units are combined.",
             "RTF recovery is optional; HTML field presence does not guarantee readable body text.",
             "Future imports using different metadata extraction may require new fingerprint logic.",
             "END OF REPORT"]
    return "\n".join(rows)+"\n"


def failure_log(exc: BaseException) -> str:
    rows = ["PST -> SQLITE METADATA RECOVERY REPORT", "="*54,
            f"Script version: {VERSION}", "Status: ABORTED; no enriched output published",
            "Original SQLite/PST: no writes requested",
            f"Failure type: {type(exc).__name__}"]
    if isinstance(exc, RecoveryError):
        rows.append(str(exc))  # RecoveryError messages are fixed, content-free strings.
    rows.append("Email text, addresses, subjects, names and paths: omitted")
    return "\n".join(rows)+"\n"


def publish_database(working: Path, target: Path):
    """Publish a complete same-filesystem copy without replacing an existing file.

    Hard-link creation is atomic and fails if target already exists, including
    a file created after CLI validation. If the filesystem cannot create links,
    fail without falling back to an overwrite or a partially copied output.
    """
    if any(os.path.lexists(path) for path in sqlite_paths(target)[1:]):
        raise RecoveryError("A sidecar already exists for the output database. Choose a fresh path.")
    os.link(working, target)
    working.unlink()


def main(argv=None, *, prog=None):
    p=argparse.ArgumentParser(prog=prog, description=__doc__)
    p.add_argument("--pst",required=True,type=Path,help="Same original PST path used for first import")
    p.add_argument("--db",required=True,type=Path,help="Existing email-funes SQLite (never modified)")
    p.add_argument("--out-db",required=True,type=Path,help="New enriched SQLite file (must not exist)")
    p.add_argument("--log",type=Path,default=Path("pst_sqlite_recovery.log"),help="Sanitized report")
    p.add_argument("--limit-matched",type=int,default=0,help="Stop after N unique matching messages (test only)")
    p.add_argument("--max-recipients",type=int,default=4096)
    p.add_argument("--progress",type=int,default=4000)
    p.add_argument("--recover-rtf",action="store_true",help="Attempt RTF text recovery; requires striprtf")
    args=p.parse_args(argv)
    pst,source=args.pst.expanduser().resolve(),args.db.expanduser().resolve()
    target,log=args.out_db.expanduser().absolute(),args.log.expanduser().absolute()
    if not pst.is_file() or pst.suffix.lower()!=".pst": p.error("--pst needs an existing .pst file")
    if not source.is_file(): p.error("--db needs an existing .sqlite file")
    try:
        require_distinct(pst,*sqlite_paths(source),*sqlite_paths(target),log)
    except FileSafetyError as exc:
        p.error(str(exc))
    if any(os.path.lexists(path) for path in sqlite_paths(target)):
        p.error("--out-db or a sidecar already exists; choose a fresh output name")
    if os.path.lexists(log): p.error("--log already exists; use a new report name")
    if args.limit_matched<0 or args.max_recipients<1 or args.progress<1: p.error("Invalid numeric argument")
    try:
        import pypff
    except ImportError:
        p.error(PST_INSTALL_HELP)
    if args.recover_rtf:
        try:
            from striprtf.striprtf import rtf_to_text  # noqa: F401
        except ImportError:
            p.error('--recover-rtf needs striprtf. From the checkout run: python -m pip install ".[rtf]"')
    try:
        private_mkdir(target.parent)
        with new_private_text(log) as report:
            return recover_to_new_file(pst, source, target, log, report, args, pypff)
    except OSError as exc:
        print(diagnostic("recovery.output", exc),file=sys.stderr)
        return 1


def recover_to_new_file(pst, source, target, log, report, args, pypff):
    fd, filename = tempfile.mkstemp(prefix=".email-funes-recovery-", suffix=".sqlite", dir=target.parent)
    os.close(fd)
    working = Path(filename)
    try:
        print("Creating a separate SQLite copy; original PST and database stay unchanged.",flush=True)
        outcome=run_recovery(pst,source,working,log,args.limit_matched,args.max_recipients,args.progress,args.recover_rtf,pypff)
        report.write(safe_log(outcome,pst))
        report.flush()
        publish_database(working,target)
        print(f"Completed. Unique rows enriched: {outcome['counts']['messages_recovered']:,}")
        print("Published the enriched database and wrote the diagnostic report.")
        return 0 if not outcome["errors"] else 2
    except KeyboardInterrupt:
        report.seek(0)
        report.truncate()
        report.write(failure_log(KeyboardInterrupt()))
        print("Interrupted; original files unchanged. Re-run with a new output path.",file=sys.stderr)
        return 130
    except Exception as exc:
        # Do not print exception text; malformed PST metadata can contain PII.
        try:
            report.seek(0)
            report.truncate()
            report.write(failure_log(exc))
        except OSError:
            print("The sanitized failure report could not be written.",file=sys.stderr)
        print(f"Recovery aborted ({type(exc).__name__}). The original files remain unchanged.",file=sys.stderr)
        if isinstance(exc, RecoveryError):
            print(str(exc),file=sys.stderr)
        return 1
    finally:
        for member in sqlite_paths(working):
            member.unlink(missing_ok=True)


if __name__=="__main__":
    raise SystemExit(main())
