# Converter reference

Run commands from the repository root. The main sequence is a source reader, SQLite, then the shared Funes exporter.

## Direct PST reader

```powershell
python -m pip install --only-binary=:all: -r requirements/pst.txt
python scripts/pst_to_sqlite_direct.py --pst "C:\MailArchives\archive.pst" --db data/email.sqlite --limit 100
```

The reader uses `pypff`, provided by the official [`libpff-python`](https://pypi.org/project/libpff-python/) package. It opens the PST directly. It does not call Outlook, send mail, or upload content.

| Option | Default | Meaning |
| --- | --- | --- |
| `--pst PATH` | Required for import | Existing PST archive |
| `--db PATH` | Required | SQLite destination |
| `--limit N` | `0` | Maximum inspected items; `0` means all |
| `--batch-size N` | `250` | Commit completed records every N inspected items |
| `--progress N` | `1000` | Print progress every N inspected items |
| `--include-junk` | Off | Include recognized Junk/Spam folders |
| `--include-deleted` | Off | Include recognized Deleted Items/Trash folders |
| `--include-header-spam` | Off | Keep messages marked by explicit spam rules |
| `--skip-bulk` | Off | Exclude List-Id, List-Unsubscribe, and bulk/list Precedence messages |
| `--max-body-chars N` | `0` | Optional storage truncation; `0` retains the full extracted body |
| `--stats` | Off | Show database counts |
| `--search QUERY` | None | Search using SQLite FTS5 syntax |
| `--results N` | `15` | Maximum search results |

Folder names are normalized for English and French names. Folder filtering also excludes their descendants. Deleted Items is a separate choice from spam; include it if deleted correspondence belongs in your project history.

The spam rules recognize `X-Spam-Flag: YES`, `X-Spam-Status: Yes`, Exchange SCL values of 5 or more, and subjects beginning `[SPAM]` or `[JUNK]`. They do not independently determine whether a message is malicious. `--skip-bulk` can remove useful technical mailing lists.

### Duplicates and resume behavior

Duplicates are matching **retained fields**: message ID, sender metadata, recipients, timestamp, subject, full extracted body, attachment names, and attachment count. Different recipients, message IDs, or body variants are preserved. Attachment bytes are not compared. The full body is fingerprinted before optional truncation.

Each source occurrence is recorded separately in `source_items`, so one email can retain several folder references. Known source entries are skipped on restart.

The direct reader's resume identifiers use the PST path, folder path, and item position. `import_sources` records file size and modification time to detect ordinary changes to that snapshot. If the PST changes, import the new snapshot into a new database. This is a metadata guard, not a full-file cryptographic hash. Do not edit a PST during extraction.

Filtered items are counted but not stored as email rows or durable exclusion records. They may be checked again on a later run. This is not continuous mailbox synchronization.

### Error handling

Individual extraction errors are recorded beside the database in `*.direct.errors.txt`. The importer continues where possible, reports missing key fields, and returns a nonzero exit code when extraction errors occur. A normal Ctrl+C commits completed records and returns 130.

For a clean import using different filter or truncation settings, create a new database. Existing rows are not rewritten by changing command-line flags.

## Outlook Classic alternative

```powershell
python -m pip install -r requirements/outlook.txt
python scripts/pst_to_sqlite.py --pst "C:\MailArchives\archive.pst" --db data/outlook-com.sqlite --limit 100
```

This Windows-only adapter uses Outlook Classic's COM object model. It needs a usable Outlook profile. The new Outlook application does not supply this interface. Open Outlook Classic and the PST manually if the adapter reports that it is not connected.

This adapter is retained for compatibility with the initial scripts. It may temporarily mount a PST in the Outlook profile; Outlook can update local PST metadata. The direct reader is the main workflow.

## SQLite to Funes

```powershell
python scripts/sqlite_to_funes.py --db data/email.sqlite --out data/funes-export
```

The exporter uses only the Python standard library. It adds a `funes_map` table to the database to persist each email's `session_id`, `turn_uuid`, and `seq`. Back up that table with the messages.

Files contain up to 2,000 turns by default; change this with `--turns-per-file N`. Files can contain several sessions, and a session can span files. Later-discovered emails append to their session's sequence even when their timestamp is older.

The output must be new or empty. The exporter stages the entire snapshot, validates essential fields, and publishes the directory after success. A failed export does not expose a partial final snapshot. Use a new folder name on the next run; unchanged turn identities let Funes deduplicate re-emitted records.

Emails with neither a sent nor a received date stay in SQLite and are skipped during Funes mapping. Invalid mapped timestamps or attachment arrays fail the export. The output currently uses `role: member` and `harness: outlook-pst`; future adapters must deliberately extend the producer metadata if they need a different facet.

## Experimental EML snapshot shortcut

The earlier script is preserved as `scripts/eml_to_funes.py`:

```powershell
python scripts/eml_to_funes.py examples/eml data/eml-funes-demo
```

It reads `.eml` files recursively, groups messages using Message-ID, References, In-Reply-To, and Outlook Thread-Index, and writes one JSONL file per thread. It uses no external package and no SQLite database. It keeps attachment filenames, excludes attachment bodies from exported text, and preserves content variants under reused message IDs.

Options: `--limit N` selects the newest N parsed messages; `--subject REGEX` filters subjects; `--max-body-chars N` optionally truncates output text. The default keeps full bodies. Messages with invalid or timezone-less Date headers are reported and skipped, never dated using filesystem modification time.

This shortcut loads the parsed snapshot into memory and does not keep a persistent thread/sequence map. A changed input set can change grouping or ordering. Use it for small, fixed snapshots; future repeatable EML imports should go through SQLite. It does not implement the PST reader's spam/folder filters. Do not index its output until the command completes successfully.

## Known limitations

- PST plain text is preferred, with HTML-to-text as a fallback. RTF-only messages may have empty bodies.
- Sender/recipient addresses read from transport headers can be missing, particularly on sent or unsent Outlook items. Missing data is not fabricated.
- Outlook conversation indexes and complete References chains can group replies; incomplete or mixed metadata can split threads. Subjects alone do not prove a shared conversation.
- Inline images can be reported as attachment filenames. Excluded documents and images cannot contribute their contents to search.
- Only email extraction is in scope; calendars, contacts, tasks, embedded-message contents, and attachment text are not a complete part of this archive model.
- Signatures, quoted history, and security banners remain in extracted bodies. No semantic rewriting or automatic reply pruning is performed.
- Tests are synthetic. Native reader compatibility and extraction completeness must be checked on the actual archive, including error logs and representative samples.
- Existing v3 databases remain readable. Existing rows, fingerprints, and session IDs are not rewritten. Use a new database for uniform application of the repository edition's deduplication and threading rules. An older populated database receives a warning when its first PST snapshot baseline is recorded.

See [Funes lifecycle behavior](funes-workflow.md) before re-indexing edited data or publishing it.
