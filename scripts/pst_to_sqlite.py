#!/usr/bin/env python3
"""Convert an Outlook Classic PST to a local, searchable SQLite archive.

Windows + Outlook Classic + pywin32 required for import. No attachments are
saved or decoded; only their filenames are recorded. Safe to restart: source
entries and duplicate messages are tracked in SQLite.

Usage:
  py -m pip install ".[outlook]"
  py pst_to_sqlite.py --pst "C:\\path\\mail.pst" --db "C:\\path\\mail.sqlite" --limit 1000
  py pst_to_sqlite.py --pst "C:\\path\\mail.pst" --db "C:\\path\\mail.sqlite"
  py pst_to_sqlite.py --db "C:\\path\\mail.sqlite" --stats
  py pst_to_sqlite.py --db "C:\\path\\mail.sqlite" --search 'maintenance OR motor'

Do not publish a database containing private or confidential messages.
"""
from __future__ import annotations

import argparse
from contextlib import closing
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
from email_funes_files import (FileSafetyError, check_writable_file, diagnostic,
                              open_import_log, open_readonly, require_distinct, sqlite_paths)


class OutlookFailure(RuntimeError):
    """Fixed diagnostic text that can safely be printed."""

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


def iter_folders(root: Any, args: argparse.Namespace, report_error):
    stack = [(root, str(safe_get(root, "Name", "PST")))]
    while stack:
        folder, path = stack.pop()
        excluded = folder_exclusion(str(safe_get(folder, "Name", "")), args)
        if excluded:
            continue
        try:
            children = folder.Folders
            for i in range(int(children.Count), 0, -1):
                child = children.Item(i)
                stack.append((child, path + "/" + str(safe_get(child, "Name", ""))))
        except Exception as exc:
            report_error("folder.children", exc, path)
        yield folder, path


def require_outlook():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise OutlookFailure('Outlook Classic and pywin32 are required. From the Windows checkout run: py -m pip install ".[outlook]"') from exc
    return pythoncom, win32com.client


def import_pst(db: sqlite3.Connection, pst: Path, args: argparse.Namespace, modules=None) -> int:
    pythoncom, client = modules if modules is not None else require_outlook()
    private = getattr(args, "private_diagnostics", False)
    errorlog = args.db.with_suffix(".private.errors.txt" if private else ".errors.txt")
    require_distinct(pst, *sqlite_paths(args.db), errorlog)
    pythoncom.CoInitialize()
    namespace = None
    newly_attached = False
    root = None
    read = inserted = duplicates = seen = filtered = errors = 0
    start = time.monotonic()
    try:
        outlook = client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        store = find_pst_store(namespace, pst)
        if store is None:
            print("Mounting PST in Outlook Classic. The application may display a prompt...", flush=True)
            namespace.AddStore(str(pst))
            newly_attached = True
            store = find_pst_store(namespace, pst)
        if store is None:
            raise OutlookFailure("Outlook could not locate the requested PST after mounting.")
        root = store.GetRootFolder()
        print("Reading the requested Outlook store.", flush=True)
        with open_import_log(errorlog) as err:
            def report_error(event, exc, location):
                nonlocal errors
                errors += 1
                err(diagnostic(event, exc, private=private, location=location))
                if errors <= 10:
                    print(diagnostic(event, exc), file=sys.stderr)

            for folder, folder_path in iter_folders(root, args, report_error):
                try:
                    items = folder.Items
                    item = items.GetFirst()
                except Exception as exc:
                    report_error("folder.items", exc, folder_path)
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
                        report_error("item.read", exc, folder_path)
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
                        report_error("folder.next", exc, folder_path)
                        break
                db.commit()
                if args.limit and read >= args.limit:
                    break
        return 2 if errors else 0
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
                print(diagnostic("store.detach", exc), file=sys.stderr)
        pythoncom.CoUninitialize()


def main(argv: Optional[list[str]] = None) -> int:
    from email_funes.adapters import legacy_pst_options

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    legacy_pst_options(parser)
    args = parser.parse_args(argv)
    return run(args, parser)


def run(args, parser):
    if args.batch_size < 1 or args.progress < 1 or args.limit < 0 or args.results < 1 or args.max_body_chars < 0:
        parser.error("limit/max-body-chars must be >=0; batch-size/progress/results must be >=1")
    args.db = args.db.expanduser().absolute()
    if args.pst:
        if os.name != "nt":
            parser.error("PST extraction via Outlook COM requires Windows. Use the .sqlite stats/search modes on other systems.")
        args.pst = args.pst.expanduser().resolve()
        if not args.pst.is_file() or args.pst.suffix.lower() != ".pst":
            parser.error("--pst must point to an existing .pst file. This script will NOT create one.")
        log = args.db.with_suffix(".private.errors.txt" if args.private_diagnostics else ".errors.txt")
        try:
            require_distinct(args.pst, *sqlite_paths(args.db), log)
            check_writable_file(log)
        except FileSafetyError as exc:
            parser.error(str(exc))
    elif not (args.stats or args.search):
        parser.error("Supply --pst to import, or --stats/--search to inspect an existing database")
    elif not args.db.is_file():
        parser.error("Database not found")

    try:
        modules = require_outlook() if args.pst else None
        with closing(open_database(args.db) if args.pst else open_readonly(args.db)) as db:
            exitcode = import_pst(db, args.pst, args, modules) if args.pst else 0
            if args.stats or args.pst:
                show_stats(db)
            if args.search:
                exitcode = search(db, args.search, args.results) or exitcode
            return exitcode
    except Exception as exc:
        print(str(exc) if isinstance(exc, (OutlookFailure, FileSafetyError))
              else diagnostic("import.failed", exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
