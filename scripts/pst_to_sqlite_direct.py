#!/usr/bin/env python3
"""Import Outlook PST directly to SQLite with pypff (no Outlook/COM required).

Windows: python -m pip install libpff-python
           python pst_to_sqlite_direct.py --pst archive.pst --db mail.sqlite --limit 10

Uses the shared SQLite schema and filters from email_funes_common.py.
Never saves or reads attachment payloads: names only. The source PST is opened
read-only. Designed for a local personal archive; don't upload confidential data.
"""
from __future__ import annotations

import argparse
import base64
import binascii
from email.header import decode_header, make_header
from email.parser import Parser
from email.utils import parseaddr
from html import unescape
from html.parser import HTMLParser
from datetime import datetime, timezone
from pathlib import Path
import json
import os
import re
import sys
import time
from typing import Any

import email_funes_common as base


class HTMLToText(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.segments = []
        self.hidden = 0

    def handle_starttag(self, tag, attrs):
        if tag in ('script', 'style'):
            self.hidden += 1
        elif tag in ('p', 'div', 'br', 'li', 'tr', 'h1', 'h2', 'h3'):
            self.segments.append('\n')

    def handle_endtag(self, tag):
        if tag in ('script', 'style') and self.hidden:
            self.hidden -= 1
        elif tag in ('p', 'div', 'li', 'tr'):
            self.segments.append('\n')

    def handle_data(self, data):
        if not self.hidden:
            self.segments.append(data)


def as_text(value: Any) -> str:
    if value is None:
        return ''
    if isinstance(value, str):
        return value
    if isinstance(value, (bytes, bytearray)):
        b = bytes(value)
        if b.startswith((b'\xff\xfe', b'\xfe\xff')):
            return b.decode('utf-16', errors='replace')
        if b.count(b'\x00') > len(b) // 5:
            return b.decode('utf-16-le', errors='replace')
        for encoding in ('utf-8', 'cp1252'):
            try:
                return b.decode(encoding)
            except UnicodeDecodeError:
                pass
        return b.decode('utf-8', errors='replace')
    return str(value)


def text_body(msg: Any) -> str:
    value = getattr(msg, 'plain_text_body', None)
    if value:
        return base.clean_body(as_text(value))
    html = getattr(msg, 'html_body', None)
    if html:
        parser = HTMLToText()
        parser.feed(as_text(html))
        return base.clean_body(unescape(''.join(parser.segments)))
    # RTF body is omitted because raw RTF is not clean searchable text and
    # this importer deliberately avoids optional third-party format parsers.
    return ''


def header_data(value: Any):
    data = as_text(value)
    return Parser().parsestr(data, headersonly=True) if data else Parser().parsestr('', headersonly=True)


def header_str(headers: Any, key: str) -> str:
    raw = headers.get(key, '') or ''
    try:
        return str(make_header(decode_header(raw))).strip()
    except (ValueError, LookupError, UnicodeError):
        return str(raw).strip()


def pff_utc_iso(value: Any) -> str | None:
    """libpff returns PST FILETIME datetimes in UTC; normalize ISO timestamp."""
    if not isinstance(value, datetime) or not (1970 <= value.year <= 2100):
        return None
    # Unlike Outlook COM timestamps, libpff's FILETIME timestamps are UTC.
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc).isoformat(timespec='seconds').replace('+00:00', 'Z')


def folder_count(folder: Any, key: str) -> int:
    # Missing optional APIs may fall back to a method; actual I/O errors must
    # remain visible so an unreadable folder is not mistaken for an empty one.
    value = getattr(folder, 'number_of_sub_' + key, None)
    if value is None:
        method = getattr(folder, 'get_number_of_sub_' + key, None)
        if not callable(method):
            raise AttributeError(f'PST reader has no {key} count API')
        value = method()
    count = int(value)
    if count < 0:
        raise ValueError(f'Invalid {key} count: {count}')
    return count


def thread_key(msg: Any, headers: Any, source_key: str) -> str:
    # The first 22 bytes of a valid Outlook conversation index identify its
    # root. Some pypff versions expose it only through the transport header.
    index = base.safe_get(msg, 'conversation_index')
    if not isinstance(index, (bytes, bytearray)) or len(index) < 22 or (len(index) - 22) % 5:
        try:
            index = base64.b64decode(''.join(header_str(headers, 'Thread-Index').split()), validate=True)
        except (ValueError, binascii.Error):
            index = b''
    if len(index) >= 22 and (len(index) - 22) % 5 == 0:
        return 'outlook-' + base.sha(bytes(index[:22]).hex())[:32]
    # The same prefix lets a root message and replies carrying its References
    # agree. Without a complete References chain, multi-hop grouping is partial.
    for field in ('References', 'In-Reply-To', 'Message-ID'):
        refs = re.findall(r'<([^<>\s]+)>', header_str(headers, field))
        if refs:
            return 'mail-' + base.sha(refs[0].casefold())[:32]
    return 'single-' + base.sha(source_key)[:32]


def register_pst_source(db, pst: Path) -> tuple[int, int]:
    """Guard positional resume keys against a replaced or edited PST snapshot."""
    path = os.path.normcase(str(pst.resolve()))
    stat = pst.stat()
    signature = (stat.st_size, stat.st_mtime_ns)
    previous = db.execute(
        'SELECT size_bytes, mtime_ns FROM import_sources WHERE source_path=?',
        (path,),
    ).fetchone()
    if previous is not None and tuple(previous) != signature:
        raise RuntimeError(
            'The PST changed since its previous import (size or modification time). '
            'Resume with the unchanged original, or import the new snapshot into '
            'a new SQLite database. Positional source IDs cannot safely resume this file.'
        )
    if previous is None:
        if db.execute('SELECT 1 FROM source_items LIMIT 1').fetchone():
            print('WARNING: recording a source baseline in a populated database. '
                  'Earlier imports cannot be verified; use the same unchanged PST.',
                  file=sys.stderr)
        db.execute('INSERT INTO import_sources(source_path,size_bytes,mtime_ns) VALUES(?,?,?)',
                   (path, *signature))
        db.commit()
    return signature


def build_message(msg: Any, pst_path: Path, folder: str, index: int, args: argparse.Namespace) -> tuple[base.Message | None, str]:
    subject = as_text(getattr(msg, 'subject', None))
    raw_headers = as_text(getattr(msg, 'transport_headers', None))
    if not args.include_header_spam and base.email_is_spam(raw_headers, subject):
        return None, 'header-spam'
    if args.skip_bulk and base.email_is_bulk(raw_headers):
        return None, 'bulk-email'
    h = header_data(raw_headers)
    full_body = text_body(msg)
    body = full_body
    if args.max_body_chars and len(body) > args.max_body_chars:
        body = body[:args.max_body_chars] + '\n[truncated by PST converter]'
    sender_name = as_text(base.safe_get(msg, 'sender_name'))
    parsed_name, sender_email = parseaddr(header_str(h, 'From') or header_str(h, 'Sender'))
    sender_name = sender_name or parsed_name
    # Current pypff bindings do not consistently expose sender/recipients
    # from all Outlook internal records; headers are best effort.
    to_list = header_str(h, 'To')
    cc_list = header_str(h, 'Cc')
    sent_utc = pff_utc_iso(base.safe_get(msg, 'client_submit_time'))
    received_utc = pff_utc_iso(base.safe_get(msg, 'delivery_time'))
    mid = base.normalize_message_id(header_str(h, 'Message-ID'))
    count = int(getattr(msg, 'number_of_attachments', 0) or 0)
    if count < 0:
        raise ValueError('Invalid attachment count')
    filenames: list[str] = []
    for a in range(count):
        try:
            att = msg.get_attachment(a)
            name = as_text(base.safe_get(att, 'long_filename'))
            filenames.append(name if name else f'[attachment {a+1}: filename unavailable]')
        except Exception:
            filenames.append(f'[attachment {a+1}: unreadable filename]')
    source_key = base.sha(os.path.normcase(str(pst_path)) + '\0' + folder + '\0' + str(index))
    fingerprint = base.message_fingerprint(
        internet_message_id=mid, sender_name=sender_name,
        sender_email=sender_email, to_list=to_list, cc_list=cc_list,
        timestamp=sent_utc or received_utc, subject=subject, full_body=full_body,
        attachment_names=filenames, attachment_count=count,
    )
    message_key = ('mid-' + base.sha(mid + '\0' + fingerprint)) if mid else 'content-' + fingerprint
    return base.Message(
        source_key=source_key, message_key=message_key,
        fingerprint=fingerprint, internet_message_id=mid,
        conversation_key=thread_key(msg, h, source_key),
        sent_utc=sent_utc, received_utc=received_utc,
        sender_name=sender_name, sender_email=sender_email,
        to_list=to_list, cc_list=cc_list, subject=subject, body=body,
        attachment_names_json=json.dumps(filenames, ensure_ascii=False),
        attachment_count=count, folder=folder,
    ), ''


def import_pst(db, pst: Path, args, pypff_module=None) -> int:
    if pypff_module is None:
        try:
            import pypff as pypff_module
        except ImportError as exc:
            raise RuntimeError('pypff is not installed. Run: python -m pip install --only-binary=:all: libpff-python; see docs/troubleshooting.md if no wheel matches your Python.') from exc
    signature = register_pst_source(db, pst)
    f = pypff_module.file()
    read = inserted = duplicate = already = filtered = errors = 0
    empty_bodies = missing_senders = missing_recipients = missing_dates = 0
    start = time.monotonic()
    errpath = args.db.with_suffix('.direct.errors.txt')
    try:
        print('Opening PST directly, without Outlook:', pst, flush=True)
        f.open(str(pst))
        root = f.get_root_folder()
        if root is None:
            raise RuntimeError('No root folder was returned by the PST reader.')
        stack = [(root, as_text(base.safe_get(root, 'name')) or 'PST')]
        with errpath.open('a', encoding='utf-8') as log:
            while stack:
                folder, path = stack.pop()
                excluded = base.folder_exclusion(path.rsplit('/',1)[-1], args)
                if excluded:
                    print('SKIP FOLDER:', path, '(', excluded, ')', flush=True)
                    continue
                print('Folder:', path, flush=True)
                try:
                    folder_children = folder_count(folder, 'folders')
                    for i in range(folder_children - 1, -1, -1):
                        child = folder.get_sub_folder(i)
                        stack.append((child, path + '/' + (as_text(base.safe_get(child,'name')) or str(i))))
                except Exception as exc:
                    errors += 1
                    log.write(f'folder={path} children_error={exc}\n')
                try:
                    n = folder_count(folder, 'messages')
                except Exception as exc:
                    errors += 1
                    log.write(f'folder={path} messages_count_error={exc}\n')
                    continue
                for i in range(n):
                    if args.limit and read >= args.limit:
                        break
                    read += 1
                    source_key = base.sha(os.path.normcase(str(pst)) + '\0' + path + '\0' + str(i))
                    if base.source_seen(db, source_key):
                        already += 1
                    else:
                        try:
                            msg = folder.get_sub_message(i)
                            record, reason = build_message(msg, pst, path, i, args)
                            if reason:
                                filtered += 1
                            elif record is not None:
                                empty_bodies += not bool(record.body.strip())
                                missing_senders += not bool(record.sender_email.strip())
                                missing_recipients += not bool(record.to_list.strip())
                                missing_dates += not bool(record.sent_utc or record.received_utc)
                                outcome = base.store_message(db, record)
                                if outcome == 'kept': inserted += 1
                                else: duplicate += 1
                        except (KeyboardInterrupt, SystemExit):
                            raise
                        except Exception as exc:
                            errors += 1
                            log.write(f'folder={path} index={i} error={type(exc).__name__}: {exc}\n')
                            if errors <= 10:
                                print('WARNING: unable to read item:', path, i, str(exc), file=sys.stderr)
                    if read % args.batch_size == 0:
                        db.commit()
                    if read % args.progress == 0:
                        elapsed = int(time.monotonic() - start)
                        print(f'inspected={read:,}, new={inserted:,}, duplicate={duplicate:,}, '
                              f'previously_imported={already:,}, filtered={filtered:,}, '
                              f'errors={errors:,}, elapsed={elapsed}s', flush=True)
                db.commit()
                if args.limit and read >= args.limit:
                    break
        final_stat = pst.stat()
        if signature != (final_stat.st_size, final_stat.st_mtime_ns):
            raise RuntimeError('PST changed during import. Completed rows were saved, '
                               'but this snapshot must be reimported into a new database.')
        return 0 if errors == 0 else 2
    except KeyboardInterrupt:
        print('\nInterrupted: SQLite changes saved. Rerun to continue.', file=sys.stderr)
        return 130
    finally:
        db.commit()
        try:
            f.close()
        except Exception:
            pass
        print(f'Run summary: inspected={read:,}; new={inserted:,}; duplicate={duplicate:,}; '
              f'previously imported={already:,}; filtered={filtered:,}; errors={errors:,}')
        if empty_bodies or missing_senders or missing_recipients or missing_dates:
            print(f'Field coverage among records read this run: empty body={empty_bodies:,}; '
                  f'missing sender address={missing_senders:,}; missing To={missing_recipients:,}; '
                  f'missing both dates={missing_dates:,}. '
                  'RTF-only bodies and absent transport headers can cause gaps.')
        if errors:
            print('Error details:', errpath)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--pst', type=Path, help='Existing PST file (no Outlook needed)')
    parser.add_argument('--db', required=True, type=Path, help='Destination SQLite file')
    parser.add_argument('--limit', type=int, default=0, help='Inspect at most N items (0=all)')
    parser.add_argument('--batch-size', type=int, default=250)
    parser.add_argument('--progress', type=int, default=1000)
    parser.add_argument('--include-junk', dest='exclude_junk', action='store_false')
    parser.add_argument('--include-deleted', dest='exclude_deleted', action='store_false')
    parser.add_argument('--include-header-spam', action='store_true')
    parser.add_argument('--skip-bulk', action='store_true', help='Optional: also omit newsletters/mailing lists')
    parser.add_argument('--max-body-chars', type=int, default=0)
    parser.add_argument('--stats', action='store_true')
    parser.add_argument('--search', metavar='FTS5_QUERY')
    parser.add_argument('--results', type=int, default=15)
    parser.set_defaults(exclude_junk=True, exclude_deleted=True)
    args = parser.parse_args(argv)
    if args.limit < 0 or args.batch_size <= 0 or args.progress <= 0 or args.max_body_chars < 0 or args.results <= 0:
        parser.error('Invalid numeric option')
    args.db = args.db.expanduser().resolve()
    if args.pst:
        args.pst = args.pst.expanduser().resolve()
        if not args.pst.is_file() or args.pst.suffix.casefold() != '.pst':
            parser.error('--pst must be an existing .pst file')
        if args.db == args.pst:
            parser.error('Database must be different from source PST')
    elif not (args.stats or args.search):
        parser.error('Specify --pst, --stats or --search')
    elif not args.db.is_file():
        parser.error('Database not found')
    db = base.open_database(args.db)
    try:
        status = import_pst(db, args.pst, args) if args.pst else 0
        if args.stats or args.pst:
            base.show_stats(db)
        if args.search:
            base.search(db, args.search, args.results)
        return status
    except Exception as exc:
        print(f'ERROR: {type(exc).__name__}: {exc}', file=sys.stderr)
        return 1
    finally:
        db.commit()
        db.close()


if __name__ == '__main__':
    raise SystemExit(main())
