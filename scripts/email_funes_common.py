"""Shared email record, SQLite storage, filters, and full-text search.

Standard library only. Source-specific readers import this module.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sqlite3
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

SCHEMA_VERSION = 1
FINGERPRINT_VERSION = 2


JUNK_NAMES = {
    "junk", "junk email", "junk e-mail", "spam", "courrier indesirable",
    "indesirables", "courrier indesirable (local)", "pourriel", "pourriels",
}


DELETED_NAMES = {
    "deleted items", "deleted", "trash", "recycle bin", "elements supprimes",
    "elements effaces", "corbeille", "eliminados",
}


def normalized_name(value: str) -> str:
    value = unicodedata.normalize("NFKD", value)
    return " ".join("".join(c for c in value if not unicodedata.combining(c)).casefold().split())


def safe_get(obj: Any, name: str, default: Any = None) -> Any:
    try:
        value = getattr(obj, name)
        return default if value is None else value
    except Exception:
        return default


def clean_body(value: str) -> str:
    """Conservative whitespace normalization: keep quoted replies intact."""
    value = value.replace("\x00", "").replace("\r\n", "\n").replace("\r", "\n")
    value = re.sub(r"[ \t]+\n", "\n", value)
    value = re.sub(r"\n{5,}", "\n\n\n", value)
    return value.strip()


def email_is_spam(headers: str, subject: str) -> bool:
    if re.search(r"(?im)^X-Spam-Flag\s*:\s*YES\b", headers):
        return True
    if re.search(r"(?im)^X-Spam-Status\s*:\s*Yes\b", headers):
        return True
    match = re.search(r"(?im)^X-MS-Exchange-Organization-SCL\s*:\s*(-?\d+)", headers)
    if match and int(match.group(1)) >= 5:
        return True
    return bool(re.match(r"\s*\[(?:spam|junk)\]", subject, flags=re.I))


def email_is_bulk(headers: str) -> bool:
    return bool(re.search(r"(?im)^(?:list-unsubscribe|list-id)\s*:", headers)
                or re.search(r"(?im)^precedence\s*:\s*(?:bulk|list)\b", headers))


def sha(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8", errors="replace")).hexdigest()


def normalize_message_id(raw: str) -> str:
    value = raw.strip().strip("<>").strip().lower()
    return value if value and len(value) < 1000 and not any(c.isspace() for c in value) else ""


def message_fingerprint(*, internet_message_id: str, sender_name: str,
                        sender_email: str, to_list: str, cc_list: str,
                        timestamp: str | None, subject: str, full_body: str,
                        attachment_names: list[str], attachment_count: int) -> str:
    """Compare full message text and retained metadata conservatively.

    IDs and recipients distinguish otherwise identical notices sent separately.
    Hash the full body before any optional storage truncation. Attachment payloads
    are deliberately excluded, so this is not a binary-equivalence guarantee.
    """
    values = [FINGERPRINT_VERSION, internet_message_id, sender_name, sender_email,
              to_list, cc_list, timestamp, subject, full_body,
              attachment_names, attachment_count]
    return sha(json.dumps(values, ensure_ascii=False, separators=(",", ":")))


@dataclass
class Message:
    source_key: str
    message_key: str
    fingerprint: str
    internet_message_id: str
    conversation_key: str
    sent_utc: Optional[str]
    received_utc: Optional[str]
    sender_name: str
    sender_email: str
    to_list: str
    cc_list: str
    subject: str
    body: str
    attachment_names_json: str
    attachment_count: int
    folder: str


def open_database(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(str(path), timeout=120)
    db.execute("PRAGMA journal_mode=WAL")
    db.execute("PRAGMA synchronous=NORMAL")
    db.execute("PRAGMA busy_timeout=120000")
    db.execute("PRAGMA foreign_keys=ON")
    db.executescript("""
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY,
            message_key TEXT NOT NULL UNIQUE,
            fingerprint TEXT NOT NULL UNIQUE,
            internet_message_id TEXT,
            conversation_key TEXT NOT NULL,
            sent_utc TEXT,
            received_utc TEXT,
            sender_name TEXT NOT NULL,
            sender_email TEXT NOT NULL,
            to_list TEXT NOT NULL,
            cc_list TEXT NOT NULL,
            subject TEXT NOT NULL,
            body TEXT NOT NULL,
            attachment_names_json TEXT NOT NULL,
            attachment_count INTEGER NOT NULL,
            first_folder TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS source_items (
            source_key TEXT PRIMARY KEY,
            message_id INTEGER REFERENCES messages(id),
            folder TEXT NOT NULL,
            disposition TEXT NOT NULL CHECK(disposition IN ('kept','duplicate'))
        );
        CREATE TABLE IF NOT EXISTS import_sources (
            source_path TEXT PRIMARY KEY,
            size_bytes INTEGER NOT NULL,
            mtime_ns INTEGER NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_messages_sent ON messages(sent_utc);
        CREATE INDEX IF NOT EXISTS idx_messages_conversation ON messages(conversation_key);
        CREATE INDEX IF NOT EXISTS idx_source_items_message ON source_items(message_id);
    """)
    try:
        db.execute("""CREATE VIRTUAL TABLE IF NOT EXISTS messages_fts USING fts5(
            subject, body, sender_name, sender_email,
            content='messages', content_rowid='id', tokenize='unicode61'
        )""")
    except sqlite3.OperationalError as exc:
        db.close()
        raise RuntimeError("Your Python's SQLite lacks FTS5 (full-text search). Install a current Python from python.org.") from exc
    db.commit()
    return db


def source_seen(db: sqlite3.Connection, source_key: str) -> bool:
    return db.execute("SELECT 1 FROM source_items WHERE source_key=?", (source_key,)).fetchone() is not None


def store_message(db: sqlite3.Connection, msg: Message) -> str:
    """Writes one message atomically. Caller batches commits for performance."""
    # An outermost SAVEPOINT is committed by RELEASE; keep a real outer
    # transaction so the caller's batch size controls disk commits.
    if not db.in_transaction:
        db.execute("BEGIN")
    db.execute("SAVEPOINT one_message")
    try:
        existing = db.execute("SELECT id FROM messages WHERE message_key=? OR fingerprint=? LIMIT 1",
                              (msg.message_key, msg.fingerprint)).fetchone()
        if existing:
            msg_id = int(existing[0])
            disposition = "duplicate"
        else:
            cur = db.execute("""INSERT INTO messages (
                message_key, fingerprint, internet_message_id, conversation_key,
                sent_utc, received_utc, sender_name, sender_email, to_list, cc_list,
                subject, body, attachment_names_json, attachment_count, first_folder
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
                msg.message_key, msg.fingerprint, msg.internet_message_id, msg.conversation_key,
                msg.sent_utc, msg.received_utc, msg.sender_name, msg.sender_email,
                msg.to_list, msg.cc_list, msg.subject, msg.body,
                msg.attachment_names_json, msg.attachment_count, msg.folder,
            ))
            msg_id = int(cur.lastrowid)
            db.execute("""INSERT INTO messages_fts(rowid, subject, body, sender_name, sender_email)
                          VALUES (?,?,?,?,?)""",
                       (msg_id, msg.subject, msg.body, msg.sender_name, msg.sender_email))
            disposition = "kept"
        db.execute("INSERT INTO source_items(source_key, message_id, folder, disposition) VALUES(?,?,?,?)",
                   (msg.source_key, msg_id, msg.folder, disposition))
        db.execute("RELEASE one_message")
        return disposition
    except Exception:
        db.execute("ROLLBACK TO one_message")
        db.execute("RELEASE one_message")
        raise


def folder_exclusion(folder_name: str, args: argparse.Namespace) -> Optional[str]:
    n = normalized_name(folder_name)
    if args.exclude_junk and n in JUNK_NAMES:
        return "junk folder"
    if args.exclude_deleted and n in DELETED_NAMES:
        return "deleted folder"
    return None


def show_stats(db: sqlite3.Connection):
    n = db.execute("SELECT COUNT(*) FROM messages").fetchone()[0]
    sources = db.execute("SELECT COUNT(*) FROM source_items").fetchone()[0]
    dup = db.execute("SELECT COUNT(*) FROM source_items WHERE disposition='duplicate'").fetchone()[0]
    attachments = db.execute("SELECT COALESCE(SUM(attachment_count), 0) FROM messages").fetchone()[0]
    size = db.execute("PRAGMA page_count").fetchone()[0] * db.execute("PRAGMA page_size").fetchone()[0]
    print(f"Unique emails:       {n:,}")
    print(f"Original PST entries:{sources:>13,}")
    print(f"Duplicate copies:    {dup:,}")
    print(f"Attachment filenames:{attachments:>13,} (no attachment contents)")
    print(f"SQLite page size:    {size/1024**2:,.1f} MiB (excluding WAL)")


def search(db: sqlite3.Connection, query: str, limit: int):
    print(f"Search: {query}\n")
    try:
        rows = db.execute("""SELECT m.sent_utc, m.sender_email, m.subject,
                              snippet(messages_fts,1,'[',']',' … ',18), m.id
                              FROM messages_fts JOIN messages m ON messages_fts.rowid=m.id
                              WHERE messages_fts MATCH ? ORDER BY rank LIMIT ?""", (query, limit))
        for date, sender, subject, snippet, msg_id in rows:
            print(f"#{msg_id} | {date or 'unknown'} | {sender}\n  {subject}\n  {snippet}\n")
    except sqlite3.OperationalError as exc:
        print(f"Search syntax error: {exc}\nUse SQLite FTS5 syntax, e.g. 'motor OR vibration' or '\"power quality\"'.", file=sys.stderr)
