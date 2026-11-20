# Adding another email format

New importers should write the shared SQLite model. That keeps Funes export, message lookup, and archive extraction independent. EML, MBOX, and Maildir are intended next formats; they are not implemented as SQLite readers yet.

## Where the code belongs

| File | Responsibility |
| --- | --- |
| `scripts/<format>_to_sqlite.py` | Parse source records, normalize metadata, provide a CLI |
| `scripts/email_funes_common.py` | Shared `Message` record, SQLite storage, filters, fingerprinting, and search |
| `scripts/sqlite_to_funes.py` | Turn the shared database into stable Funes records |
| `tests/test_<format>.py` | Synthetic extraction and failure-case tests |
| `examples/` | Small invented inputs using reserved example domains |

## Importer contract

1. Stream messages when the source permits; a large mailbox should not need to fit in RAM.
2. Build a stable `source_key` for an occurrence. Prefer a source-native identifier. Positional identifiers need an immutable source snapshot guard.
3. Preserve message dates and normalize to UTC. Leave an unknown date missing; do not substitute the extraction date.
4. Extract sender, recipients, subject, body, and attachment **filenames**. Skip attachment content, including nested attached emails.
5. Use `message_fingerprint()` on the full body and retained metadata before any optional storage truncation. Preserve distinct versions and recipients.
6. Assign a conservative, namespaced `conversation_key`; use a separate session when threading evidence is absent.
7. Use `source_seen()` before repeating work and `store_message()` for atomic message/source/FTS writes. Commit in bounded batches.
8. Report errors and missing fields. A failed parse must not silently become a complete empty record.

The shared writer is used like this inside an importer (the source-specific reader must supply a `Message`):

```python
from pathlib import Path
from email_funes_common import open_database, source_seen, store_message

def import_records(records, destination: Path):
    db = open_database(destination)
    try:
        for index, message in enumerate(records, start=1):
            if not source_seen(db, message.source_key):
                store_message(db, message)
            if index % 250 == 0:
                db.commit()
        db.commit()
    finally:
        db.close()
```

Use the complete [SQLite schema](sqlite-schema.md) for required fields. Keep source-specific code out of the shared SQLite module.

## Thread and producer identity

Use stable message IDs and thread roots. Subject text alone is insufficient: two unrelated messages can have the same subject. Do not renumber exported turns or recalculate existing thread IDs after a memory has been indexed.

The current shared exporter labels records `outlook-pst`. When adding another source, extend its producer selection deliberately (for example a persisted source column and a compatible migration), rather than silently labeling all formats as PST. Keep existing producers and `funes_map` identities stable. The experimental direct EML script is historical code, not the template for incremental synchronization.

## Tests that matter

- Plain text, HTML fallback, non-ASCII subjects, and timezone conversion.
- The same email in two folders, distinct recipients, reused Message-IDs, and different bodies with identical truncated prefixes.
- Attached text files and `message/rfc822` payloads never appearing in exported bodies.
- Resume after interruption, a changed source snapshot, and corrupt records.
- Missing metadata reported honestly and dates never invented.
- A successful synthetic SQLite → JSONL export with stable sequence assignments.

Use invented content. Add the reader to the README's capability table only when it exists and has a documented command. Keep optional dependencies in a separate requirements file so existing users do not have to install every parser.
