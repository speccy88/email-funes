#!/usr/bin/env python3
"""Convert an EML snapshot directly to Funes JSONL (experimental).

Usage: python scripts/eml_to_funes.py /path/to/eml /path/to/funes-out

Produces one .funes.jsonl file per email thread. Uses Message-ID,
References, In-Reply-To and Outlook Thread-Index for conservative threading.
No network access or non-stdlib dependencies. Attachments are not indexed;
only their filenames are retained as text metadata.

This is a *snapshot* importer. For repeated/incremental imports, preserve
thread IDs and sequence assignments in an external database (e.g. SQLite).
"""
from __future__ import annotations

import argparse
import base64
from collections import defaultdict
from datetime import datetime, timezone
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
import hashlib
from html.parser import HTMLParser
import json
from pathlib import Path
import re
import sys

from email_funes_files import FileSafetyError, diagnostic, new_private_text, staged_directory

MSG_ID = re.compile(r"<([^<>\s]+)>")


def sha(value: str | bytes) -> str:
    if isinstance(value, str):
        value = value.encode("utf-8", "replace")
    return hashlib.sha256(value).hexdigest()


def get_ids(header: str) -> list[str]:
    matches = MSG_ID.findall(header)
    if not matches:
        matches = [s.strip(" <>,\t\r\n") for s in header.split() if "@" in s]
    return [s for s in matches if s]


class HtmlToText(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.hidden = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self.hidden += 1
        elif tag in ("p", "br", "div", "li", "tr", "h1", "h2", "h3"):
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self.hidden = max(0, self.hidden - 1)
        elif tag in ("p", "div", "li", "tr"):
            self.parts.append("\n")

    def handle_data(self, data):
        if not self.hidden:
            self.parts.append(data)


def normalize_html(raw: str) -> str:
    parser = HtmlToText()
    parser.feed(raw)
    result = "".join(parser.parts)
    return re.sub(r"\n{3,}", "\n\n", result).strip()


def visible_parts(msg):
    """Walk body parts and attachment metadata without entering attachment subtrees."""
    if msg.get_content_disposition() == "attachment" or msg.get_filename():
        yield msg
        return
    if msg.get_content_type() == "message/rfc822":
        return
    if msg.is_multipart():
        for child in msg.iter_parts():
            yield from visible_parts(child)
    else:
        yield msg


def message_text(msg) -> str:
    plain, html = [], []
    for part in visible_parts(msg):
        if part.get_content_disposition() == "attachment" or part.get_filename():
            continue
        typ = part.get_content_type()
        if typ not in ("text/plain", "text/html"):
            continue
        try:
            value = part.get_content()
        except (LookupError, UnicodeError, ValueError):
            payload = part.get_payload(decode=True) or b""
            value = payload.decode("utf-8", "replace")
        if typ == "text/plain":
            plain.append(value)
        else:
            html.append(value)
    return "\n".join(plain).strip() if plain else normalize_html("\n".join(html))


def utc_timestamp(value: str) -> str:
    try:
        d = parsedate_to_datetime(value)
        if d.tzinfo is None:
            raise ValueError("Date header has no timezone")
    except (TypeError, ValueError, OverflowError, IndexError) as exc:
        raise ValueError("Missing or invalid Date header; file modification time is not a message date") from exc
    return d.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def outlook_root(value: str) -> str | None:
    try:
        raw = base64.b64decode("".join(value.split()), validate=True)
    except (ValueError, base64.binascii.Error):
        return None
    # The first 22 bytes of a Microsoft conversation index identify the thread.
    return raw[:22].hex() if len(raw) >= 22 and (len(raw) - 22) % 5 == 0 else None


class UnionFind:
    def __init__(self):
        self.parent = {}

    def find(self, x: str) -> str:
        if x not in self.parent:
            self.parent[x] = x
        while x != self.parent[x]:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def join(self, a: str, b: str):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[max(ra, rb)] = min(ra, rb)


def read_message(path: Path, root: Path, max_chars: int) -> dict | None:
    raw = path.read_bytes()
    msg = BytesParser(policy=policy.default).parsebytes(raw)
    mid = get_ids(str(msg.get("Message-ID", "")))
    id_value = mid[0] if mid else f"sha256-{sha(raw)}"
    refs = get_ids(str(msg.get("References", "")))
    in_reply_to = get_ids(str(msg.get("In-Reply-To", "")))
    thread_index = outlook_root(str(msg.get("Thread-Index", "")))
    attachments = [part.get_filename() for part in visible_parts(msg) if part.get_filename()]
    folder = str(path.parent.relative_to(root))
    subject = str(msg.get("Subject", ""))
    sender = str(msg.get("From", ""))
    recipients = str(msg.get("To", ""))
    cc = str(msg.get("Cc", ""))
    date = utc_timestamp(str(msg.get("Date", "")))
    body = message_text(msg)
    # Keep original message provenance in indexed text because the JSONL schema is strict.
    text = "\n".join((
        f"Subject: {subject}", f"From: {sender}", f"To: {recipients}",
        f"Cc: {cc}", f"Date (UTC): {date}", f"PST folder: {folder}",
        f"Message-ID: {id_value}",
        f"Attachments: {', '.join(attachments) if attachments else '(none)'}",
        "", body[:max_chars] if max_chars else body,
        "\n[Body truncated by importer]" if max_chars and len(body) > max_chars else "",
    ))
    return {
        "id": id_value, "refs": refs + in_reply_to, "ti": thread_index,
        "date": date, "text": text, "source": str(path),
        "variant": sha(json.dumps([id_value, date, sender, recipients, cc, subject, body, attachments], ensure_ascii=False)),
    }


def convert(src: Path, dst: Path, limit: int | None, subject_filter: str | None,
            max_chars: int) -> int:
    paths = sorted(p for p in src.rglob("*") if p.is_file() and p.suffix.lower() == ".eml")
    if not paths:
        raise FileSafetyError("No .eml files found under the source directory.")
    if src.resolve() == dst.resolve():
        raise FileSafetyError("Source and output directories must differ.")
    search = re.compile(subject_filter, re.IGNORECASE) if subject_filter else None
    data = []
    failed = 0
    for path in paths:
        try:
            item = read_message(path, src, max_chars)
            if search and not search.search(item["text"].split("\n", 1)[0]):
                continue
            data.append(item)
        except Exception as exc:
            failed += 1
            print(diagnostic("eml.parse", exc), file=sys.stderr)
    data.sort(key=lambda x: (x["date"], x["id"], x["source"]))
    if limit:
        data = data[-limit:]  # pilot uses the newest N messages
    # Preserve content variants even when a sender reused the Message-ID.
    uniq = {}
    for item in data:
        uniq.setdefault(item["variant"], item)
    data = list(uniq.values())

    uf = UnionFind()
    for item in data:
        own = f"mid:{item['id']}"
        uf.find(own)
        for ref in item["refs"]:
            uf.join(own, f"mid:{ref}")
        if item["ti"]:
            uf.join(own, f"thread-index:{item['ti']}")

    groups = defaultdict(list)
    for item in data:
        groups[uf.find(f"mid:{item['id']}")].append(item)

    written = 0
    with staged_directory(dst) as staging:
        for root_key, messages in groups.items():
            # Snapshot identities, not an incremental SQLite mapping.
            session_id = "pst-" + sha(root_key)[:24]
            with new_private_text(staging / f"{session_id}.tmp") as f:
                for seq, item in enumerate(messages):
                    turn = {
                        "format": 1, "session_id": session_id,
                        "turn_uuid": "mail-" + item["variant"][:24],
                        "seq": seq, "ts": item["date"],
                        "role": "member", "harness": "outlook-pst",
                        "blocks": [{"block_type": "text", "text": item["text"]}],
                    }
                    f.write(json.dumps(turn, ensure_ascii=False, separators=(",", ":")) + "\n")
                    written += 1
        for path in staging.glob("*.tmp"):
            path.rename(path.with_suffix(".funes.jsonl"))
    print(f"Read {len(paths)} EML files; exported {written} distinct emails across {len(groups)} threads.")
    if failed:
        print(f"WARNING: {failed} messages could not be parsed; review and correct before archival.")
    return 2 if failed else 0


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("eml_dir", type=Path, help="Folder containing extracted .eml messages")
    parser.add_argument("output_dir", type=Path, help="New or empty output folder")
    parser.add_argument("--limit", type=int, default=None, help="Process only the newest N emails for a pilot")
    parser.add_argument("--subject", default=None, help="Optional regex to filter by email subject")
    parser.add_argument("--max-body-chars", type=int, default=0, help="0 keeps the full body (default)")
    args = parser.parse_args()
    if not args.eml_dir.is_dir():
        parser.error("eml_dir must be an existing directory")
    if args.limit is not None and args.limit < 1:
        parser.error("--limit must be at least 1")
    if args.max_body_chars < 0:
        parser.error("--max-body-chars must be at least 0")
    try:
        return convert(args.eml_dir, args.output_dir, args.limit, args.subject, args.max_body_chars)
    except FileSafetyError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    except Exception as exc:
        print(diagnostic("eml.failed", exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
