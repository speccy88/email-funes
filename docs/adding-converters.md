# Adding another email format

New importers should write the shared SQLite model. That keeps Funes export, message lookup, and archive extraction independent. EML, MBOX, and Maildir are intended next formats; they are not implemented as SQLite readers yet.

## Where the code belongs

| File | Responsibility |
| --- | --- |
| `scripts/email_funes/adapters.py` | Dispatch `email-funes import --format` and expose supported readers |
| `scripts/<format>_to_sqlite.py` | Parse source-specific records and normalize metadata |
| `scripts/email_funes_common.py` | Shared `Message` record, SQLite storage, filters, fingerprinting, and search |
| `scripts/sqlite_to_funes.py` | Apply a stable selection/rendering policy and export Funes records |
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
9. If the source mixes email with other objects, provide explicit classification and document which categories the exporter selects. Preserve typed recipients and provenance when available. Do not silently change an existing export policy.

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

When multiple archives feed one database, `source_items` records each occurrence, while the fingerprint identifies matching extracted messages. Identically named mailboxes in different directories can contain different mail. The direct PST reader keeps its historical path/folder/position keys and guards positions with a size and modification-time baseline. Keep the same unchanged original at that path. These attributes are continuity checks, not content authentication.

For a new adapter, namespace its source occurrence IDs by format and archive identity. Freeze the source snapshot when offsets or indexes participate in the key. Keep stable message IDs and thread roots. Subject lines can repeat across unrelated conversations. Extend the dispatch table only after the reader, tests and installation instructions are ready. Keep native dependencies inside the selected reader.

The shared exporter's existing producer label is `outlook-pst`. A future adapter needs a deliberate persisted producer/provenance policy for new and duplicate records, including messages shared across formats. Add a versioned transactional migration with old-database and mapping tests if new columns are needed. Existing PST assignments and rendered content must remain stable. A recovery or rendering change can require a rebuild of the derived memory.

Sequence numbers belong to the database's persistent `funes_map`. Import all related snapshots into one database before exporting a combined memory. Independent databases can assign the same sequence number to different messages in a shared conversation. Combine independent exports only after verifying agreement on producer, session, sequence, turn UUID and rendered content, or index them in separate memories.

## Next adapters

| Planned adapter | First implementation step |
| --- | --- |
| EML | Share MIME parsing with the experimental snapshot exporter where appropriate. Write normalized messages and per-file provenance through the SQLite writer. Include repeated Message-IDs and duplicate filenames in tests. |
| MBOX | Stream messages from each fixed archive, using an immutable snapshot identity to guard per-message offsets on resume. Test rewrites and truncated records. |
| Maildir | Track each mailbox and message occurrence across `new/` and `cur/`. Test renames, flag suffixes and identical filenames in separate mailboxes. |

The experimental `scripts/eml_to_funes.py` still exports a standalone JSONL snapshot. The public CLI's SQLite import currently accepts PST.

## Tests that matter

- Plain text, HTML fallback, non-ASCII subjects, and timezone conversion.
- The same email in two folders, distinct recipients, reused Message-IDs, and different bodies with identical truncated prefixes.
- Attached text files and `message/rfc822` payloads never appearing in exported bodies.
- Resume after interruption, a changed source snapshot, and corrupt records.
- Missing metadata reported honestly and dates never invented.
- A successful synthetic SQLite → JSONL export with stable sequence assignments.

Use invented content. Add the reader to the README's capability table only when it exists and has a documented command. Keep optional dependencies in a separate requirements file so existing users do not have to install every parser.
