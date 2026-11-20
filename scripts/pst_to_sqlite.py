#!/usr/bin/env python3
"""Convert an Outlook Classic PST to a local, searchable SQLite archive.

Windows + Outlook Classic + pywin32 required for import. No attachments are
saved or decoded; only their filenames are recorded. Safe to restart: source
entries and duplicate messages are tracked in SQLite.

Usage:
  py -m pip install pywin32
  py pst_to_sqlite.py --pst "C:\\path\\mail.pst" --db "C:\\path\\mail.sqlite" --limit 1000
  py pst_to_sqlite.py --pst "C:\\path\\mail.pst" --db "C:\\path\\mail.sqlite"
  py pst_to_sqlite.py --db "C:\\path\\mail.sqlite" --stats
  py pst_to_sqlite.py --db "C:\\path\\mail.sqlite" --search 'maintenance OR motor'

Do not publish a database containing private or confidential messages.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sqlite3
import sys
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

from email_funes_common import (
    Message, message_fingerprint, normalize_message_id, clean_body, email_is_bulk, email_is_spam, folder_exclusion, normalized_name, open_database, safe_get, search, sha, show_stats, source_seen, store_message
)

MAIL_ITEM_CLASS = 43
MAPI_MESSAGE_ID = (
    "http://schemas.microsoft.com/mapi/proptag/0x1035001F",
    "http://schemas.microsoft.com/mapi/proptag/0x1035001E",
)
MAPI_HEADERS = (
    "http://schemas.microsoft.com/mapi/proptag/0x007D001F",
    "http://schemas.microsoft.com/mapi/proptag/0x007D001E",
)
MAPI_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"


def mapi_prop(obj: Any, tags: tuple[str, ...] | str) -> str:
    try:
        accessor = obj.PropertyAccessor
    except Exception:
        return ""
    if isinstance(tags, str):
        tags = (tags,)
    for tag in tags:
        try:
            value = accessor.GetProperty(tag)
            if value:
                return str(value)
        except Exception:
            pass
    return ""


def to_utc_iso(value: Any) -> Optional[str]:
    """Outlook's COM dates typically arrive as local-time datetimes."""
    if value is None:
        return None
    try:
        # Strip COM date wrappers into a standard Python datetime.
        dt = datetime(int(value.year), int(value.month), int(value.day),
                      int(value.hour), int(value.minute), int(value.second),
                      int(getattr(value, "microsecond", 0)))
        if not (1970 <= dt.year <= 2100):
            return None
        if getattr(value, "tzinfo", None) is not None:
            dt = dt.replace(tzinfo=value.tzinfo)
        # Naive datetime.astimezone() treats the value as machine local time.
        return dt.astimezone(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    except (TypeError, ValueError, AttributeError, OverflowError):
        return None


def canonical_subject(value: str) -> str:
    value = re.sub(r"^(?:(?:re|fw|fwd|tr|réf|sv|aw)\s*:\s*)+", "", value.strip(), flags=re.I)
    return " ".join(value.casefold().split())


def build_message(item: Any, path: str, folder: str, skip_headers: bool,
                  skip_bulk: bool, max_body_chars: int) -> tuple[Optional[Message], str]:
    """Read a MailItem into plain Python types (without saving attachments)."""
    entry_id = str(safe_get(item, "EntryID", "") or "")
    # If EntryID is absent, source identity is determined later from content.
    source_key = sha(os.path.normcase(path) + "\0" + entry_id) if entry_id else ""

    subject = str(safe_get(item, "Subject", "") or "")
    headers = "" if skip_headers and not skip_bulk else mapi_prop(item, MAPI_HEADERS)
    if not skip_headers and email_is_spam(headers, subject):
        return None, "header-spam"
    if skip_bulk and email_is_bulk(headers):
        return None, "bulk-email"

    # Property errors must reach the per-message error log, not create empty
    # "successful" imports that cannot be retried.
    full_body = clean_body(str(getattr(item, "Body", "") or ""))
    body = full_body
    if max_body_chars and len(body) > max_body_chars:
        body = body[:max_body_chars] + "\n[truncated by PST converter]"

    filenames: list[str] = []
    attachments = safe_get(item, "Attachments")
    n = int(safe_get(attachments, "Count", 0) or 0)
    for i in range(1, n + 1):
        try:
            attachment = attachments.Item(i)
            filename = str(safe_get(attachment, "FileName", "") or "")
            filenames.append(filename or f"[attachment {i}: filename unavailable]")
        except Exception:
            # Attachment metadata may be corrupt. Do not save or decode payloads.
            filenames.append("[unreadable attachment filename]")

    sent_utc = to_utc_iso(safe_get(item, "SentOn"))
    received_utc = to_utc_iso(safe_get(item, "ReceivedTime"))
    sender_name = str(safe_get(item, "SenderName", "") or "")
    sender_email = mapi_prop(item, MAPI_SENDER_SMTP) or str(safe_get(item, "SenderEmailAddress", "") or "")
    to_list = str(safe_get(item, "To", "") or "")
    cc_list = str(safe_get(item, "CC", "") or "")
    internet_id = normalize_message_id(mapi_prop(item, MAPI_MESSAGE_ID))
    conversation_id = str(safe_get(item, "ConversationID", "") or "").strip()
    fingerprint = message_fingerprint(
        internet_message_id=internet_id, sender_name=sender_name,
        sender_email=sender_email, to_list=to_list, cc_list=cc_list,
        timestamp=sent_utc or received_utc, subject=subject, full_body=full_body,
        attachment_names=filenames, attachment_count=n,
    )
    if not entry_id:
        source_key = sha(os.path.normcase(path) + "\0" + folder + "\0" + fingerprint)
    # Fallback is one Funes session per message, avoiding unrelated subject merges.
    conversation_key = "outlook-" + sha(conversation_id)[:32] if conversation_id else "single-" + sha(source_key)[:32]
    # Retain different body variants even if a source reuses the same Message-ID.
    message_key = ("mid-" + sha(internet_id + "\0" + fingerprint)) if internet_id else ("content-" + fingerprint)
    msg = Message(
        source_key=source_key, message_key=message_key, fingerprint=fingerprint,
        internet_message_id=internet_id, conversation_key=conversation_key,
        sent_utc=sent_utc, received_utc=received_utc,
        sender_name=sender_name, sender_email=sender_email,
        to_list=to_list, cc_list=cc_list, subject=subject, body=body,
        attachment_names_json=json.dumps(filenames, ensure_ascii=False),
        attachment_count=n, folder=folder,
    )
    return msg, ""


def find_pst_store(namespace: Any, pst: Path) -> Any:
    target = os.path.normcase(os.path.normpath(str(pst)))
    stores = namespace.Stores
    for i in range(1, int(stores.Count) + 1):
        s = stores.Item(i)
        candidate = str(safe_get(s, "FilePath", "") or "")
        if candidate and os.path.normcase(os.path.normpath(candidate)) == target:
            return s
    return None


def iter_folders(root: Any, args: argparse.Namespace):
    stack = [(root, str(safe_get(root, "Name", "PST")))]
    while stack:
        folder, path = stack.pop()
        excluded = folder_exclusion(str(safe_get(folder, "Name", "")), args)
        if excluded:
            print(f"SKIP FOLDER ({excluded}): {path}", flush=True)
            continue
        try:
            children = folder.Folders
            for i in range(int(children.Count), 0, -1):
                child = children.Item(i)
                stack.append((child, path + "/" + str(safe_get(child, "Name", ""))))
        except Exception as exc:
            print(f"WARNING unable to list subfolders for {path}: {exc}", file=sys.stderr)
        yield folder, path


def import_pst(db: sqlite3.Connection, pst: Path, args: argparse.Namespace) -> int:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32 is required: py -m pip install pywin32") from exc

    pythoncom.CoInitialize()
    namespace = None
    newly_attached = False
    root = None
    read = inserted = duplicates = seen = filtered = errors = 0
    start = time.monotonic()
    errorlog = args.db.with_suffix(".errors.txt")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        store = find_pst_store(namespace, pst)
        if store is None:
            print("Mounting PST in Outlook Classic. The application may display a prompt...", flush=True)
            namespace.AddStore(str(pst))
            newly_attached = True
            store = find_pst_store(namespace, pst)
        if store is None:
            raise RuntimeError("PST could not be located in Outlook's data stores after mounting.")
        root = store.GetRootFolder()
        print("Importing ONLY this PST:", pst, flush=True)
        with errorlog.open("a", encoding="utf-8") as err:
            for folder, folder_path in iter_folders(root, args):
                print("Folder:", folder_path, flush=True)
                try:
                    items = folder.Items
                    item = items.GetFirst()
                except Exception as exc:
                    errors += 1
                    err.write(f"folder={folder_path} error={exc}\n")
                    continue
                while item is not None:
                    if args.limit and read >= args.limit:
                        break
                    try:
                        if int(safe_get(item, "Class", 0) or 0) != MAIL_ITEM_CLASS:
                            item = items.GetNext()
                            continue
                        read += 1
                        # Check an inexpensive stable source identifier before reading body/attachments.
                        entry_id = str(safe_get(item, "EntryID", "") or "")
                        if entry_id and source_seen(db, sha(os.path.normcase(str(pst)) + "\0" + entry_id)):
                            seen += 1
                        else:
                            msg, reason = build_message(
                                item, str(pst), folder_path, args.include_header_spam,
                                args.skip_bulk, args.max_body_chars,
                            )
                            if reason:
                                filtered += 1
                            elif msg is not None and source_seen(db, msg.source_key):
                                seen += 1
                            elif msg is not None:
                                disposition = store_message(db, msg)
                                if disposition == "kept":
                                    inserted += 1
                                else:
                                    duplicates += 1
                    except Exception as exc:
                        errors += 1
                        err.write(f"folder={folder_path} error={type(exc).__name__}: {exc}\n")
                        if errors <= 10:
                            print(f"WARNING: one message skipped in {folder_path}: {exc}", file=sys.stderr)
                    finally:
                        if read and read % args.batch_size == 0:
                            db.commit()
                        if read and read % args.progress == 0:
                            elapsed = int(time.monotonic() - start)
                            print(f"read={read:,}, stored={inserted:,}, duplicates={duplicates:,}, "
                                  f"already seen={seen:,}, spam/bulk={filtered:,}, errors={errors:,}, "
                                  f"elapsed={elapsed}s", flush=True)
                    try:
                        item = items.GetNext()
                    except Exception as exc:
                        errors += 1
                        err.write(f"folder={folder_path} iteration_error={exc}\n")
                        print(f"WARNING: folder iteration stopped: {folder_path}: {exc}", file=sys.stderr)
                        break
                db.commit()
                if args.limit and read >= args.limit:
                    break
        return errors
    except KeyboardInterrupt:
        print("\nInterrupted. Committing completed messages; rerun the same command to resume.", file=sys.stderr)
        db.commit()
        return 130
    finally:
        db.commit()
        print(f"Completed this run: inspected={read:,}; new={inserted:,}; duplicate={duplicates:,}; "
              f"previously imported={seen:,}; filtered={filtered:,}; errors={errors:,}")
        if newly_attached and root is not None and namespace is not None:
            try:
                namespace.RemoveStore(root)
                print("PST detached from Outlook profile (file is still on disk).")
            except Exception as exc:
                print(f"NOTE: Could not detach PST automatically: {exc}", file=sys.stderr)
        pythoncom.CoUninitialize()


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--pst", type=Path, help="Existing .pst file (requires Windows and Outlook Classic)")
    parser.add_argument("--db", required=True, type=Path, help="Destination SQLite file")
    parser.add_argument("--limit", type=int, default=0, help="Trial run: inspect at most N email items (0=all)")
    parser.add_argument("--batch-size", type=int, default=250, help="Commit every N emails (default 250)")
    parser.add_argument("--progress", type=int, default=1000, help="Report progress every N emails")
    parser.add_argument("--include-junk", dest="exclude_junk", action="store_false", help="Include Junk/Spam folders")
    parser.add_argument("--include-deleted", dest="exclude_deleted", action="store_false", help="Include Deleted Items folders")
    parser.add_argument("--include-header-spam", action="store_true", help="Do not filter messages with clear spam headers")
    parser.add_argument("--skip-bulk", action="store_true", help="Also exclude mailing lists/newsletters (off by default)")
    parser.add_argument("--max-body-chars", type=int, default=0, help="Optional limit per email body (0=unlimited)")
    parser.add_argument("--stats", action="store_true", help="Show DB statistics (no Outlook required)")
    parser.add_argument("--search", metavar="FTS5_QUERY", help="Search the DB (no Outlook required)")
    parser.add_argument("--results", type=int, default=15, help="Maximum search hits (default 15)")
    parser.set_defaults(exclude_junk=True, exclude_deleted=True)
    args = parser.parse_args(argv)
    if args.batch_size < 1 or args.progress < 1 or args.limit < 0 or args.results < 1 or args.max_body_chars < 0:
        parser.error("limit/max-body-chars must be >=0; batch-size/progress/results must be >=1")
    args.db = args.db.expanduser().resolve()
    if args.pst:
        if os.name != "nt":
            parser.error("PST extraction via Outlook COM requires Windows. Use the .sqlite stats/search modes on other systems.")
        args.pst = args.pst.expanduser().resolve()
        if not args.pst.is_file() or args.pst.suffix.lower() != ".pst":
            parser.error("--pst must point to an existing .pst file. This script will NOT create one.")
        if args.pst == args.db:
            parser.error("PST and SQLite paths must differ")
    elif not (args.stats or args.search):
        parser.error("Supply --pst to import, or --stats/--search to inspect an existing database")
    elif not args.db.is_file():
        parser.error("Database not found")

    db = open_database(args.db)
    try:
        exitcode = 0
        if args.pst:
            try:
                exitcode = import_pst(db, args.pst, args)
            except Exception as exc:
                print(f"ERROR: {exc}", file=sys.stderr)
                return 1
        if args.stats or args.pst:
            show_stats(db)
        if args.search:
            search(db, args.search, args.results)
        return exitcode
    finally:
        db.commit()
        db.close()


if __name__ == "__main__":
    raise SystemExit(main())
