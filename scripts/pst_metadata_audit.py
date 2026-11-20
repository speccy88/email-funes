#!/usr/bin/env python3
"""Read-only Outlook PST metadata diagnostic for email-funes.

Requires Python 3.10+ and `libpff-python` (import name `pypff`).

Never prints email addresses, message subjects, bodies, attachment names,
message identifiers, raw folder names, or local paths to the diagnostic log.
Does not write to the PST or the optional SQLite database.

Typical use (PowerShell):
  python pst_metadata_audit.py --pst "C:\\Archives\\archive.pst" \
      --db "C:\\Archives\\outlook_archive.sqlite" --out pst_metadata_audit.log

For a full pass: add --full. For a quick pilot: --per-folder 20.
"""
from __future__ import annotations

import argparse
from collections import Counter
from datetime import datetime, timezone
from email.parser import Parser
import hashlib
import hmac
import os
from pathlib import Path
import secrets
import sqlite3
import sys
import time
import unicodedata
from typing import Any

VERSION = "1.1.0"
MAX_RECORD_SETS = 16
MAX_ENTRIES_PER_SET = 2048
MAX_RECIPIENTS_PER_ITEM = 512

# MAPI property IDs (16-bit, excluding property type). Keep sensitive values
# only in process memory to measure their recoverability, never in the report.
PROPERTIES = {
    0x001A: "message_class",
    0x0037: "subject",
    0x0065: "sent_representing_address",
    0x0070: "conversation_topic",
    0x007D: "transport_headers",
    0x0C1A: "sender_name",
    0x0C1E: "sender_address_type",
    0x0C1F: "sender_email_address",
    0x0E02: "display_bcc",
    0x0E03: "display_cc",
    0x0E04: "display_to",
    0x1000: "plain_body_present",
    0x1009: "rtf_body_present",
    0x1013: "html_body_present",
    0x1035: "internet_message_id",
    0x5D01: "sender_smtp_address",
    0x5D02: "representing_smtp_address",
}
RECIPIENT_PROPERTIES = {
    0x0C15: "recipient_type",         # 1=TO, 2=CC, 3=BCC
    0x3001: "display_name",
    0x3002: "address_type",           # SMTP or EX, etc.
    0x3003: "email_address",
    0x39FE: "smtp_address",
}
STRING_VALUE_TYPES = {0x001E, 0x001F}  # MAPI PT_STRING8 / PT_UNICODE
TEXT_MESSAGE_FIELDS = {k for k in PROPERTIES if k not in (0x1000, 0x1009, 0x1013, 0x0070, 0x007D)}

# These property IDs are numbers, not addresses or message contents.
PROPERTY_LABELS = {**PROPERTIES, **{k: f"recipient_{v}" for k, v in RECIPIENT_PROPERTIES.items()}}


def sanitized_error(exc: BaseException) -> str:
    """Never include exception messages: they may include mailbox content."""
    return type(exc).__name__


def text_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        if value.startswith((b"\xff\xfe", b"\xfe\xff")):
            return value.decode("utf-16", "replace")
        if value.count(b"\x00") > len(value) // 5:
            return value.decode("utf-16-le", "replace")
        for encoding in ("utf-8", "cp1252"):
            try:
                return value.decode(encoding)
            except UnicodeError:
                continue
        return value.decode("utf-8", "replace")
    return str(value)


def try_attr(obj: Any, name: str, errors: Counter | None = None) -> Any:
    try:
        return getattr(obj, name)
    except AttributeError:
        return None
    except Exception as exc:
        if errors is not None:
            errors[f"attribute.{name}.{sanitized_error(exc)}"] += 1
        return None


def get_number(obj: Any, attr: str, method: str, errors: Counter) -> int:
    value = try_attr(obj, attr, errors)
    if value is None:
        fun = try_attr(obj, method, errors)
        if not callable(fun):
            raise AttributeError(f"{attr} / {method} unavailable")
        value = fun()
    result = int(value)
    if result < 0:
        raise ValueError("negative count")
    return result


def property_id(raw: Any) -> int | None:
    try:
        n = int(raw)
        return (n >> 16) & 0xFFFF if n > 0xFFFF else n
    except (TypeError, ValueError):
        return None


def extract_properties(obj: Any, wanted: set[int], errors: Counter,
                       source: str) -> tuple[set[int], dict[int, str], dict[int, int]]:
    """Enumerate MAPI field IDs and decode only a small, explicit allowlist.

    The function never reads any entry's raw bytes or the body/attachment payload.
    """
    found: set[int] = set()
    decoded: dict[int, str] = {}
    value_types: dict[int, int] = {}
    try:
        # Recipient rows are pypff.record_set objects (not pypff.item), while
        # messages contain one or more record_set objects. Support both shapes.
        if callable(try_attr(obj, "get_entry", errors)) and not callable(try_attr(obj, "get_record_set", errors)):
            record_sets = [obj]
        else:
            nsets = get_number(obj, "number_of_record_sets", "get_number_of_record_sets", errors)
            if nsets > MAX_RECORD_SETS:
                errors[f"{source}.record_sets_capped"] += 1
            record_sets = [obj.get_record_set(si) for si in range(min(nsets, MAX_RECORD_SETS))]
        for rs in record_sets:
            try:
                if rs is None:
                    continue
                total = get_number(rs, "number_of_entries", "get_number_of_entries", errors)
                if total > MAX_ENTRIES_PER_SET:
                    errors[f"{source}.entries_capped"] += 1
                for ei in range(min(total, MAX_ENTRIES_PER_SET)):
                    try:
                        entry = rs.get_entry(ei)
                        if entry is None:
                            continue
                        pid = property_id(try_attr(entry, "entry_type", errors))
                        if pid is None:
                            continue
                        found.add(pid)
                        vt = try_attr(entry, "value_type", errors)
                        if vt is not None:
                            try:
                                value_types[pid] = int(vt)
                            except (TypeError, ValueError):
                                pass
                        if pid == 0x0C15 and pid in wanted:
                            if hasattr(entry, "get_data_as_integer"):
                                try:
                                    number = entry.get_data_as_integer()
                                    if number is not None:
                                        decoded[pid] = str(int(number))
                                except Exception as exc:
                                    errors[f"{source}.recip_type.{sanitized_error(exc)}"] += 1
                        elif pid in wanted and value_types.get(pid) in STRING_VALUE_TYPES:
                            if pid not in decoded or not decoded[pid]:
                                try:
                                    # This is the only MAPI content read: targeted metadata.
                                    s = entry.get_data_as_string()
                                    if s:
                                        decoded[pid] = text_value(s).strip()
                                except Exception as exc:
                                    errors[f"{source}.string.{hex(pid)}.{sanitized_error(exc)}"] += 1
                    except Exception as exc:
                        errors[f"{source}.entry.{sanitized_error(exc)}"] += 1
            except Exception as exc:
                errors[f"{source}.set.{sanitized_error(exc)}"] += 1
    except Exception as exc:
        errors[f"{source}.record_sets.{sanitized_error(exc)}"] += 1
    return found, decoded, value_types


def normalize_folder(name: str) -> str:
    s = unicodedata.normalize("NFKD", name).casefold()
    return "".join(c for c in s if not unicodedata.combining(c))


def folder_category(path: str) -> str:
    p = normalize_folder(path)
    parts = [x.strip() for x in p.split("/")]
    for x in reversed(parts):
        if any(z in x for z in ("junk", "spam", "pourriel", "indesirable")):
            return "Junk / Spam"
        if any(z in x for z in ("deleted items", "trash", "elements supprimes", "corbeille")):
            return "Deleted"
        if any(z in x for z in ("sent items", "sent mail", "elements envoyes", "messages envoyes", "envoyes")) or x == "sent":
            return "Sent"
        if any(z in x for z in ("inbox", "boite de reception", "reception")):
            return "Inbox"
        if any(z in x for z in ("calendar", "calendrier")):
            return "Calendar"
        if any(z in x for z in ("contacts", "carnet d'adresse")):
            return "Contacts"
        if any(z in x for z in ("tasks", "taches", "tâches")):
            return "Tasks"
        if any(z in x for z in ("drafts", "brouillons")):
            return "Drafts"
        if any(z in x for z in ("outbox", "boite d'envoi")):
            return "Outbox"
    return "Other / project folder"


def item_class_bucket(value: str) -> str:
    x = value.casefold().strip()
    if x.startswith("ipm.note") or x.startswith("report.ipm.note"):
        return "Email / IPM.Note"
    if x.startswith("ipm.schedule.meeting"):
        return "Meeting request"
    if x.startswith("ipm.appointment"):
        return "Appointment"
    if x.startswith("ipm.contact") or x.startswith("ipm.distlist"):
        return "Contact / distribution list"
    if x.startswith("ipm.task"):
        return "Task"
    if x.startswith("ipm.stickynote"):
        return "Note"
    return "Unknown / other"  # Do not disclose arbitrary custom message-class strings.


def record_header_presence(value: str) -> dict[str, bool]:
    if not value:
        return {k: False for k in ("header_from", "header_to", "header_cc", "header_message_id", "header_date", "header_references", "header_thread_index")}
    try:
        h = Parser().parsestr(value, headersonly=True)
        return {
            "header_from": bool(h.get("From") or h.get("Sender")),
            "header_to": bool(h.get("To")),
            "header_cc": bool(h.get("Cc")),
            "header_message_id": bool(h.get("Message-ID")),
            "header_date": bool(h.get("Date")),
            "header_references": bool(h.get("References") or h.get("In-Reply-To")),
            "header_thread_index": bool(h.get("Thread-Index")),
        }
    except Exception:
        return {k: False for k in ("header_from", "header_to", "header_cc", "header_message_id", "header_date", "header_references", "header_thread_index")}


def recipient_probe(msg: Any, errors: Counter) -> dict[str, bool | int]:
    result: dict[str, bool | int] = {"recipient_table_available": False, "recipients_count": 0,
        "recipient_any_address": False, "recipient_smtp": False,
        "recipient_exchange": False, "recipient_to": False, "recipient_cc": False,
        "recipient_bcc": False, "recipient_untyped": False,
        "recipient_property_ids": ()}
    recipients = try_attr(msg, "recipients", errors)
    if recipients is None:
        return result
    recipient_property_ids: set[int] = set()
    try:
        n = get_number(recipients, "number_of_recipients", "get_number_of_recipients", errors)
        result["recipient_table_available"] = True
        result["recipients_count"] = n
        if n > MAX_RECIPIENTS_PER_ITEM:
            errors["recipients.capped"] += 1
        for index in range(min(n, MAX_RECIPIENTS_PER_ITEM)):
            try:
                rec = recipients.get_recipient(index)
                if rec is None:
                    continue
                present, decoded, vtypes = extract_properties(rec, set(RECIPIENT_PROPERTIES), errors, "recipient")
                recipient_property_ids.update(present)
                # A recipient can use SMTP, Exchange X.500, or other address forms.
                smtp = decoded.get(0x39FE, "")
                address = decoded.get(0x3003, "")
                addr_type = decoded.get(0x3002, "").upper()
                has_addr = bool(smtp or address)
                if has_addr:
                    result["recipient_any_address"] = True
                    if smtp and "@" in smtp or (address and "@" in address):
                        result["recipient_smtp"] = True
                    if addr_type == "EX" or address.casefold().startswith("/o="):
                        result["recipient_exchange"] = True
                    typ = decoded.get(0x0C15, "")
                    if typ == "1": result["recipient_to"] = True
                    elif typ == "2": result["recipient_cc"] = True
                    elif typ == "3": result["recipient_bcc"] = True
                    else: result["recipient_untyped"] = True
                if (0x3003 in present or 0x39FE in present) and not has_addr:
                    errors["recipient.address_present_but_unreadable"] += 1
            except Exception as exc:
                errors[f"recipients.item.{sanitized_error(exc)}"] += 1
    except Exception as exc:
        errors[f"recipients.table.{sanitized_error(exc)}"] += 1
    result["recipient_property_ids"] = tuple(sorted(recipient_property_ids))
    return result


def message_probe(msg: Any, errors: Counter) -> dict[str, bool | str | int]:
    # Avoid reading body data or attachment data. Only metadata is inspected.
    transport = text_value(try_attr(msg, "transport_headers", errors))
    headers = record_header_presence(transport)
    present, vals, vtypes = extract_properties(msg, TEXT_MESSAGE_FIELDS | {0x001A}, errors, "message")
    attr_sender = text_value(try_attr(msg, "sender_email_address", errors)) or text_value(try_attr(msg, "sender_smtp_address", errors))
    attr_class = text_value(try_attr(msg, "message_class", errors))
    message_class = vals.get(0x001A, "") or attr_class
    recipients = recipient_probe(msg, errors)
    sender_raw = vals.get(0x5D01) or vals.get(0x0C1F) or vals.get(0x5D02) or vals.get(0x0065) or attr_sender
    smtp_like = bool(sender_raw and "@" in sender_raw)
    result: dict[str, bool | str | int] = {
        "class": item_class_bucket(message_class),
        "message_class_known": bool(message_class),
        "transport_headers_present": bool(transport),
        "sender_name_native": bool(try_attr(msg, "sender_name", errors)),
        "sender_attr_available": bool(attr_sender),
        "sender_raw_address": bool(sender_raw),
        "sender_raw_smtp_like": smtp_like,
        "sender_raw_exchange_like": bool(sender_raw and (sender_raw.casefold().startswith("/o=") or sender_raw.casefold().startswith("ex:"))),
        "mapi_sender_addr": bool(vals.get(0x0C1F)),
        "mapi_sender_smtp": bool(vals.get(0x5D01)),
        "mapi_repr_addr": bool(vals.get(0x0065)),
        "mapi_repr_smtp": bool(vals.get(0x5D02)),
        "display_to": bool(vals.get(0x0E04)),
        "display_cc": bool(vals.get(0x0E03)),
        "mapi_internet_id": bool(vals.get(0x1035)),
        "mapi_subject": bool(vals.get(0x0037)),
        "native_subject": bool(try_attr(msg, "subject", errors)),
        "conversation_index": bool(try_attr(msg,"conversation_index",errors)),
        "mapi_plain_field": 0x1000 in present,
        "mapi_html_field": 0x1013 in present,
        "mapi_rtf_field": 0x1009 in present,
        "mapi_rtf_only": 0x1009 in present and 0x1000 not in present and 0x1013 not in present,
        **headers,
        **recipients,
    }
    result["sender_recoverable"] = bool(result["header_from"] or result["sender_raw_address"])
    result["to_recoverable"] = bool(result["header_to"] or result["display_to"] or result["recipient_to"])
    result["to_possibly_recoverable"] = bool(result["to_recoverable"] or result["recipient_untyped"])
    result["cc_recoverable"] = bool(result["header_cc"] or result["display_cc"] or result["recipient_cc"])
    result["email_likely"] = result["class"] == "Email / IPM.Note" or (
        not result["message_class_known"] and (result["transport_headers_present"] or result["sender_raw_address"]))
    result["known_message_property_count"] = len(present & set(PROPERTIES))
    result["message_property_ids"] = tuple(sorted(present))
    # No real content, values, subject, personal identifiers, or folder paths returned.
    return result


def sampled_indices(count: int, per_folder: int) -> list[int] | range:
    if per_folder <= 0 or count <= per_folder:
        return range(count)
    return sorted({min(count - 1, ((2 * i + 1) * count) // (2 * per_folder)) for i in range(per_folder)})


def safe_folder_id(secret: bytes, name: str) -> str:
    return hmac.new(secret, name.encode("utf-8", "replace"), hashlib.sha256).hexdigest()[:10]


def open_old_db(path: Path | None, errors: Counter) -> sqlite3.Connection | None:
    if path is None:
        return None
    if not path.is_file():
        errors["db.missing_file"] += 1
        return None
    try:
        conn = sqlite3.connect(path.resolve().as_uri() + "?mode=ro", uri=True)
        conn.execute("SELECT m.sender_email,m.to_list,m.subject,m.body FROM messages m LIMIT 1")
        conn.execute("SELECT source_key,message_id FROM source_items LIMIT 1")
        return conn
    except Exception as exc:
        errors[f"db.open.{sanitized_error(exc)}"] += 1
        return None


def old_record_flags(db: sqlite3.Connection, pst: Path, path: str, index: int) -> tuple[bool, bool, bool, bool] | None:
    old_hash = hashlib.sha256((os.path.normcase(str(pst)) + "\0" + path + "\0" + str(index)).encode("utf-8", "replace")).hexdigest()
    row = db.execute("""
        SELECT (m.sender_email IS NOT NULL AND TRIM(m.sender_email) <> ''),
               (m.to_list IS NOT NULL AND TRIM(m.to_list) <> ''),
               (m.subject IS NOT NULL AND TRIM(m.subject) <> ''),
               (m.body IS NOT NULL AND TRIM(m.body) <> '')
        FROM source_items s INNER JOIN messages m ON m.id=s.message_id
        WHERE s.source_key = ?
        """, (old_hash,)).fetchone()
    return tuple(bool(x) for x in row) if row else None


def pct(n: int, total: int) -> str:
    return f"{n} / {total} ({100*n/total:.1f}%)" if total else "0 / 0 (n/a)"


class Stats:
    def __init__(self):
        self.errors = Counter()
        self.fields = Counter()
        self.message_property_ids = Counter()
        self.recipient_property_ids = Counter()
        self.type_counts = Counter()
        self.folder_categories = Counter()
        self.folder_infos: list[tuple[str, str, int, int, int, int]] = []
        self.sampled = 0
        self.total_slots = 0
        self.folders = 0
        self.db_compared = 0
        self.db_nomatch = 0
        self.db_missing_sender = 0
        self.db_missing_to = 0
        self.db_missing_subject = 0
        self.db_subject_potential = 0
        self.db_missing_body = 0
        self.db_body_plain_field = 0
        self.db_body_html_field = 0
        self.db_body_rtf_field = 0
        self.db_sender_potential = 0
        self.db_sender_smtp_potential = 0
        self.db_to_potential = 0
        self.db_to_smtp_potential = 0
        self.db_to_untyped_potential = 0
        self.db_missing_sender_by_type = Counter()
        self.db_to_recover_by_type = Counter()
        self.db_email_matched = 0
        self.db_email_missing_sender = 0
        self.db_email_missing_to = 0
        self.db_email_sender_potential = 0
        self.db_email_to_potential = 0
        self.fatal = None


def scan(pst: Path, db: sqlite3.Connection | None, per_folder: int,
         progress: int, pypff_module: Any | None = None) -> Stats:
    stats = Stats()
    if pypff_module is None:
        import pypff as pypff_module
    file = pypff_module.file()
    secret = secrets.token_bytes(32)
    try:
        file.open(str(pst))
        root = file.get_root_folder()
        if root is None:
            raise RuntimeError("PST reader returned no root folder")
        root_name = text_value(try_attr(root, "name", stats.errors)) or "PST"
        stack = [(root, root_name)]
        while stack:
            folder, path = stack.pop()
            stats.folders += 1
            safe_id = safe_folder_id(secret, path)
            category = folder_category(path)
            category_count = 0
            inspected = 0
            missing_sender = 0
            missing_to = 0
            try:
                children = get_number(folder, "number_of_sub_folders", "get_number_of_sub_folders", stats.errors)
                for i in range(children - 1, -1, -1):
                    try:
                        child = folder.get_sub_folder(i)
                        name = text_value(try_attr(child, "name", stats.errors)) or "[unnamed]"
                        stack.append((child, path + "/" + name))
                    except Exception as exc:
                        stats.errors[f"folder.child.{sanitized_error(exc)}"] += 1
            except Exception as exc:
                stats.errors[f"folder.child_count.{sanitized_error(exc)}"] += 1
            try:
                count = get_number(folder, "number_of_sub_messages", "get_number_of_sub_messages", stats.errors)
            except Exception as exc:
                stats.errors[f"folder.message_count.{sanitized_error(exc)}"] += 1
                continue
            stats.total_slots += count
            samples = sampled_indices(count, per_folder)
            category_count += count
            for idx in samples:
                inspected += 1
                stats.sampled += 1
                stats.folder_categories[category] += 1
                try:
                    msg = folder.get_sub_message(idx)
                    r = message_probe(msg, stats.errors)
                    stats.type_counts[str(r["class"])] += 1
                    stats.message_property_ids.update(r.pop("message_property_ids", ()))
                    stats.recipient_property_ids.update(r.pop("recipient_property_ids", ()))
                    for k, v in r.items():
                        if k not in {"class", "recipients_count", "known_message_property_count"} and v:
                            stats.fields[k] += 1
                    if db is not None:
                        original = old_record_flags(db, pst, path, idx)
                        if original is None:
                            stats.db_nomatch += 1
                        else:
                            stats.db_compared += 1
                            old_sender, old_to, old_subject, old_body = original
                            if not old_subject:
                                stats.db_missing_subject += 1
                                if r["mapi_subject"] or r["native_subject"]:
                                    stats.db_subject_potential += 1
                            if not old_body:
                                stats.db_missing_body += 1
                                if r["mapi_plain_field"]: stats.db_body_plain_field += 1
                                if r["mapi_html_field"]: stats.db_body_html_field += 1
                                if r["mapi_rtf_field"]: stats.db_body_rtf_field += 1
                            if r["class"] == "Email / IPM.Note":
                                stats.db_email_matched += 1
                                if not old_sender:
                                    stats.db_email_missing_sender += 1
                                    if r["sender_raw_address"] or r["header_from"]:
                                        stats.db_email_sender_potential += 1
                                if not old_to:
                                    stats.db_email_missing_to += 1
                                    if r["to_recoverable"]:
                                        stats.db_email_to_potential += 1
                            if not old_sender:
                                missing_sender += 1
                                stats.db_missing_sender += 1
                                stats.db_missing_sender_by_type[str(r["class"])] += 1
                                if r["sender_raw_address"] or r["header_from"]:
                                    stats.db_sender_potential += 1
                                if r["sender_raw_smtp_like"]:
                                    stats.db_sender_smtp_potential += 1
                            if not old_to:
                                missing_to += 1
                                stats.db_missing_to += 1
                                if r["to_recoverable"]:
                                    stats.db_to_potential += 1
                                    stats.db_to_recover_by_type[str(r["class"])] += 1
                                if r["recipient_smtp"]:
                                    stats.db_to_smtp_potential += 1
                                if not r["to_recoverable"] and r["recipient_untyped"]:
                                    stats.db_to_untyped_potential += 1
                except Exception as exc:
                    stats.errors[f"message.probe.{sanitized_error(exc)}"] += 1
                if progress and stats.sampled % progress == 0:
                    print(f"Inspected {stats.sampled:,} items in {stats.folders:,} folders...", flush=True)
            stats.folder_infos.append((safe_id, category, count, inspected, missing_sender, missing_to))
    except Exception as exc:
        stats.fatal = sanitized_error(exc)
    finally:
        try:
            file.close()
        except Exception as exc:
            stats.errors[f"file.close.{sanitized_error(exc)}"] += 1
    return stats


def report(stats: Stats, pst: Path, db_supplied: bool, per_folder: int,
           reader_version: str, max_folder_rows: int) -> str:
    n = stats.sampled
    lines = [
        "PST METADATA DIAGNOSTIC (privacy-safe summary)",
        "=" * 65,
        f"Diagnostic script version: {VERSION}",
        f"Reader: libpff / pypff ({reader_version})",
        "PST location and filename: [REDACTED]",
        f"PST size: {pst.stat().st_size / (1024**3):.2f} GiB",
        "SQLite location and filename: [REDACTED]",
        f"SQLite comparison requested: {'yes' if db_supplied else 'no'}",
        f"Scan mode: {'FULL' if per_folder == 0 else f'spread sample, up to {per_folder} per folder'}",
        "Data protection: no raw subjects, addresses, bodies, names, or attachments logged",
        "", "[SCAN COVERAGE]",
        f"Folders inspected: {stats.folders:,}",
        f"Total message slots reported by PST folders: {stats.total_slots:,}",
        f"Messages actually inspected: {stats.sampled:,}",
        f"Sample fraction: {pct(stats.sampled, stats.total_slots)}",
        "Note: message slots can include appointments, contacts, and tasks.",
        "", "[MESSAGE CLASSES - SAMPLED]",
    ]
    for label, cnt in stats.type_counts.most_common():
        lines.append(f"{label:32s} {pct(cnt, n)}")
    lines.extend(["", "[PST METADATA AVAILABILITY - SAMPLED]"])
    keys = [
        ("Headers: any transport headers", "transport_headers_present"),
        ("Headers: From/Sender present", "header_from"),
        ("Headers: To present", "header_to"),
        ("Headers: Cc present", "header_cc"),
        ("Headers: Internet Message-ID present", "header_message_id"),
        ("MAPI: subject property decoded", "mapi_subject"),
        ("Headers: References / In-Reply-To", "header_references"),
        ("Headers: Thread-Index", "header_thread_index"),
        ("libpff: native sender name", "sender_name_native"),
        ("libpff: native sender address attr", "sender_attr_available"),
        ("MAPI: sender email property decoded", "mapi_sender_addr"),
        ("MAPI: sender SMTP property decoded", "mapi_sender_smtp"),
        ("MAPI: representing sender email", "mapi_repr_addr"),
        ("MAPI: representing sender SMTP", "mapi_repr_smtp"),
        ("Sender address found in non-header source", "sender_raw_address"),
        ("Sender address looks like SMTP (@)", "sender_raw_smtp_like"),
        ("MAPI: display-To property decoded", "display_to"),
        ("MAPI: display-Cc property decoded", "display_cc"),
        ("Recipient table accessible", "recipient_table_available"),
        ("Recipient table: address available", "recipient_any_address"),
        ("Recipient table: SMTP-like address", "recipient_smtp"),
        ("Recipient table: Exchange-style address", "recipient_exchange"),
        ("Recipient table: typed To address", "recipient_to"),
        ("Recipient table: typed Cc address", "recipient_cc"),
        ("Recipient table: typed Bcc address", "recipient_bcc"),
        ("Recipient table: address without known type", "recipient_untyped"),
        ("Message class available (MAPI or attr)", "message_class_known"),
        ("PST has plain-text body MAPI field", "mapi_plain_field"),
        ("PST has HTML body MAPI field", "mapi_html_field"),
        ("PST has RTF body MAPI field", "mapi_rtf_field"),
        ("PST has RTF field but neither plain nor HTML", "mapi_rtf_only"),
    ]
    for label, key in keys:
        lines.append(f"{label:49s} {pct(stats.fields[key], n)}")
    lines.extend(["", "[COMMON MAPI PROPERTY IDS - COUNTS ONLY]",
                  "Properties are reported by numeric ID, never by their stored values."])
    for pid, count in stats.message_property_ids.most_common(28):
        label = PROPERTIES.get(pid, "unmapped")
        lines.append(f"  Message 0x{pid:04X}  {label:27s} {pct(count, n)}")
    for pid, count in stats.recipient_property_ids.most_common(15):
        label = RECIPIENT_PROPERTIES.get(pid, "unmapped")
        lines.append(f"  Recipient 0x{pid:04X}  {label:25s} {count} sampled messages")
    lines.extend(["", "[SQLITE BASELINE COMPARISON - MATCHED SOURCE POSITIONS]"])
    if not db_supplied:
        lines.append("Not requested. Supply --db to compare missing fields.")
    else:
        x = stats.db_compared
        lines.append(f"Records matched by original PST path/folder/item position: {pct(x, n)}")
        lines.append(f"Records not matched (moved PST, filtered, changed, etc.): {stats.db_nomatch:,}")
        lines.append(f"Matched SQLite rows missing sender address: {pct(stats.db_missing_sender, x)}")
        lines.append(f"  Among missing: raw MAPI/native or header sender value exists: {pct(stats.db_sender_potential, stats.db_missing_sender)}")
        lines.append(f"  Among missing: sender value looks SMTP-like: {pct(stats.db_sender_smtp_potential, stats.db_missing_sender)}")
        lines.append(f"Matched SQLite rows missing To: {pct(stats.db_missing_to, x)}")
        lines.append(f"  Among missing: To header/display/typed table available: {pct(stats.db_to_potential, stats.db_missing_to)}")
        lines.append(f"  Among missing: any SMTP-like recipient (type may be unknown): {pct(stats.db_to_smtp_potential, stats.db_missing_to)}")
        lines.append(f"  Among missing: only untyped recipient candidate: {pct(stats.db_to_untyped_potential, stats.db_missing_to)}")
        lines.append(f"Matched SQLite rows missing subject: {pct(stats.db_missing_subject, x)}")
        lines.append(f"  Among missing: raw subject available: {pct(stats.db_subject_potential, stats.db_missing_subject)}")
        lines.append(f"Matched SQLite rows missing body: {pct(stats.db_missing_body, x)}")
        lines.append(f"  Among missing: plain-text field exists: {pct(stats.db_body_plain_field, stats.db_missing_body)}")
        lines.append(f"  Among missing: HTML field exists: {pct(stats.db_body_html_field, stats.db_missing_body)}")
        lines.append(f"  Among missing: RTF field exists: {pct(stats.db_body_rtf_field, stats.db_missing_body)}")
        lines.append("Known IPM.Note messages only (excludes classified contacts/calendar/tasks):")
        lines.append(f"  Matched email-class rows: {stats.db_email_matched:,}")
        lines.append(f"  Email rows missing sender: {pct(stats.db_email_missing_sender, stats.db_email_matched)}")
        lines.append(f"  Missing sender with recovery candidate: {pct(stats.db_email_sender_potential, stats.db_email_missing_sender)}")
        lines.append(f"  Email rows missing To: {pct(stats.db_email_missing_to, stats.db_email_matched)}")
        lines.append(f"  Missing To with recovery candidate: {pct(stats.db_email_to_potential, stats.db_email_missing_to)}")
        if x == 0:
            lines.append("WARNING: 0 matches. Confirm --pst points to the same original path as the SQLite import.")
        lines.append("Potential does not guarantee that an SMTP address is available or correct.")
        lines.append("The optional SQLite file was opened read-only; no records were changed.")
        lines.append("Missing sender breakdown by message class:")
        for category, count in stats.db_missing_sender_by_type.most_common(12):
            lines.append(f"  {category:31s} {count:,}")
    lines.extend(["", "[FOLDER CATEGORIES - SAMPLED]", "Folder categories are guessed from names but names are not logged."])
    for category, count in stats.folder_categories.most_common():
        lines.append(f"  {category:26s} {count:,}")
    lines.extend(["", "[LARGEST FOLDERS - ANONYMIZED]",
        "IDs are randomized per scan and cannot be used to recover folder names.",
        "ID           category                    slots  sampled  missing sender  missing To"])
    for folder_id, category, slots, sampled, miss_sender, miss_to in sorted(stats.folder_infos, key=lambda x: x[2], reverse=True)[:max_folder_rows]:
        lines.append(f"{folder_id:12s} {category[:27]:27s} {slots:6,d}  {sampled:7,d} {miss_sender:14,d} {miss_to:11,d}")
    lines.extend(["", "[SANITIZED ERRORS - COUNTS ONLY]"])
    if stats.errors:
        for label, count in stats.errors.most_common(35):
            lines.append(f"  {label}: {count:,}")
    else:
        lines.append("  No API errors recorded.")
    if stats.fatal:
        lines.append(f"FATAL: {stats.fatal} (exception text intentionally excluded)")
    lines.extend(["", "[INTERPRETATION]",
        "- Existing converter relies on transport headers, which may be absent on sent messages.",
        "- MAPI sender properties and recipient tables can provide alternative metadata.",
        "- Exchange/X.500 addresses are not necessarily valid SMTP addresses; don't invent conversions.",
        "- Non-email message classes should be routed separately before Funes indexing.",
        "- Body field presence does not guarantee successful text extraction or decoding.",
        "- Sampled results describe inspected items, not necessarily the entire archive.",
        "- This log contains structural counts, not personal email content.",
        "END OF REPORT"])
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--pst", required=True, type=Path, help="PST archive, opened read-only")
    parser.add_argument("--db", type=Path, help="Existing email-funes SQLite database (optional, read-only)")
    parser.add_argument("--out", type=Path, default=Path("pst_metadata_audit.log"), help="Output diagnostic .log file")
    parser.add_argument("--per-folder", type=int, default=40, help="Distributed samples per folder; 0 for full scan (default: 40)")
    parser.add_argument("--full", action="store_true", help="Examine all items instead of a sample")
    parser.add_argument("--progress", type=int, default=2000, help="Print progress every N examined items")
    parser.add_argument("--max-report-folders", type=int, default=20, help="Maximum anonymized folder summary rows")
    args = parser.parse_args(argv)
    args.pst = args.pst.expanduser().resolve()
    if not args.pst.is_file() or args.pst.suffix.casefold() != ".pst":
        parser.error("--pst must point to an existing .pst file")
    if args.db is not None:
        args.db = args.db.expanduser().resolve()
    args.out = args.out.expanduser().resolve()
    if args.out == args.pst or (args.db and args.out == args.db):
        parser.error("The diagnostic log must not overwrite the PST or SQLite file")
    if args.out.exists() and any(path and path.exists() and args.out.samefile(path)
                                for path in (args.pst, args.db)):
        parser.error("The diagnostic log must not be a hard link to the PST or SQLite file")
    if args.per_folder < 0 or args.progress < 1 or args.max_report_folders < 0:
        parser.error("Invalid numeric option")
    if args.full:
        args.per_folder = 0
    errors = Counter()
    db = open_old_db(args.db, errors)
    try:
        import pypff
    except ImportError:
        print("ERROR: module pypff missing. Install: python -m pip install --only-binary=:all: libpff-python",file=sys.stderr)
        if db: db.close()
        return 2
    try:
        version = text_value(pypff.get_version()) if callable(getattr(pypff, "get_version", None)) else "unknown"
    except Exception:
        version = "unknown"
    print("Reading PST metadata. No emails or attachments will be exported.", flush=True)
    started = time.monotonic()
    try:
        stats = scan(args.pst, db, args.per_folder, args.progress, pypff)
        stats.errors.update(errors)
        output = report(stats, args.pst, args.db is not None, args.per_folder, version, args.max_report_folders)
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(output, encoding="utf-8")
        elapsed = time.monotonic() - started
        print(f"Finished: {stats.sampled:,} sampled of {stats.total_slots:,} message slots in {stats.folders:,} folders; elapsed {elapsed:.1f} s.")
        print(f"Wrote privacy-safe log: {args.out}")
        print("Send this .log file here; there is no need to share the PST.")
        if stats.fatal:
            print("Scan stopped with a sanitized error; see FATAL in log.",file=sys.stderr)
            return 1
        return 0
    finally:
        if db is not None:
            db.close()


if __name__ == "__main__":
    raise SystemExit(main())
