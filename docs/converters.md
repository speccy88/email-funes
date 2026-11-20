# Converter reference

Run commands from the repository root. The main sequence is a source reader, SQLite, optional metadata recovery, then the shared Funes exporter.

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

## PST metadata audit and recovery

Use `pst_metadata_audit.py` to measure native property coverage and `pst_sqlite_recover.py` to enrich a fresh copy of an existing import. These commands share the `pypff` dependency. Recovery can also decode RTF text with optional `striprtf`.

```powershell
python scripts/pst_metadata_audit.py --pst "C:\MailArchives\archive.pst" --db data/email.sqlite --out data/audit.log --full
python scripts/pst_sqlite_recover.py --pst "C:\MailArchives\archive.pst" --db data/email.sqlite --out-db data/email_enriched.sqlite --log data/recovery.log --recover-rtf
```

The original PST path and immutable snapshot are part of source matching. Recovery does not add filtered source items or rewrite existing message identities. It adds item classification, structured recipients, recovered fields, and provenance to the copy. Its aggregate log contains no message text or addresses.

See the [metadata recovery guide](metadata-recovery.md) for dependencies, flags, pilot runs, source checks, and limitations.

## SQLite to Funes

```powershell
python scripts/sqlite_to_funes.py --db data/email_enriched.sqlite --dry-run
python scripts/sqlite_to_funes.py --db data/email_enriched.sqlite --out data/funes-export
```

The exporter uses only the Python standard library. Preview is read-only: it reports category, selection, missing-date, and existing-mapping counts without producing JSONL or changing SQLite. It checks policy compatibility but does not fully validate message content or timestamp syntax.

| Option | Default | Meaning |
| --- | --- | --- |
| `--db PATH` | Required | Existing email-funes SQLite database |
| `--out PATH` | Required unless `--dry-run` | New or empty JSONL destination directory |
| `--dry-run` | Off | Read-only selection and policy preview |
| `--include-meetings` | Off | Include meeting messages alongside email |
| `--include-all` | Off | Include all retained item categories; mutually exclusive with `--include-meetings` |
| `--include-bcc` | Off | Include available Bcc recipients in the rendered turn text |
| `--turns-per-file N` | `2000` | Maximum turns per output file |

### Record selection and recipients

When recovery classification exists, the default selects the `Email / IPM.Note` category. `--include-meetings` also selects `Meeting request`, which covers meeting-message subclasses such as invitations and responses. Appointments, contacts, tasks, notes, and unclassified records require `--include-all`.

Without recovery classification, the exporter retains the earlier all-record behavior and prints a notice. Use recovery to obtain an email-only selection based on native message classes.

Normalized recipient rows take precedence over the original flattened To/Cc text. If a recipient type has no normalized rows, the exporter uses the original field where available. Existing but empty typed rows do not fall back to an older display list. It retains unresolved display names rather than inventing an address. Bcc is opt-in because its inclusion changes the searchable content.

Only the selected records receive new Funes mappings. Records with neither a sent nor a received date stay in SQLite and are skipped; timestamps are never invented. Records with an empty body can still provide searchable subjects, participants, dates, and attachment filenames.

### Stable export state

A normal export adds two tables to SQLite:

- `funes_map` keeps each selected record's `session_id`, `turn_uuid`, and dense `seq`.
- `funes_export_policy` keeps the selection and recipient-rendering choices, including Bcc inclusion.

Back up these tables together with the messages. Later-discovered emails append to their session's sequence even when their timestamp is older. Files can contain several sessions, and a session can span files.

The exporter rejects a different policy or an existing mapping that conflicts with the selected records. It does not renumber already assigned turns. An older mapping without a policy can be adopted only when its legacy rendering and selection behavior are preserved. An enriched database with an earlier legacy mapping needs a deliberate fresh export state.

If a policy change is necessary, preserve the existing database/export and start from a separate copy of the archive taken **before its first mapping**. Use a separate derived Funes memory for a rebuild. Do not delete mapping rows in a published workflow to force the command through: that can shift sequences or retain old indexed content under reused identities.

### Output and validation

The output must be new or empty. The exporter stages the entire snapshot, validates essential fields, and publishes the directory after success. A failed export does not expose a partial final snapshot. Use a new folder name on the next run; unchanged turn identities let Funes deduplicate re-emitted records.

Invalid mapped timestamps, malformed attachment filename arrays, or non-dense sequences fail the export. It uses `role: member` and `harness: outlook-pst`; email metadata stays inside one text block. The `role` field is not a sender address, and no undocumented top-level JSON fields are added.

Run `funes index --check` on the completed files before indexing. See the [Funes lifecycle guide](funes-workflow.md) for the consequences of changing already-indexed content.

## Experimental EML snapshot shortcut

The earlier script is preserved as `scripts/eml_to_funes.py`:

```powershell
python scripts/eml_to_funes.py examples/eml data/eml-funes-demo
```

It reads `.eml` files recursively, groups messages using Message-ID, References, In-Reply-To, and Outlook Thread-Index, and writes one JSONL file per thread. It uses no external package and no SQLite database. It keeps attachment filenames, excludes attachment bodies from exported text, and preserves content variants under reused message IDs.

Options: `--limit N` selects the newest N parsed messages; `--subject REGEX` filters subjects; `--max-body-chars N` optionally truncates output text. The default keeps full bodies. Messages with invalid or timezone-less Date headers are reported and skipped, never dated using filesystem modification time.

This shortcut loads the parsed snapshot into memory and does not keep a persistent thread/sequence map. A changed input set can change grouping or ordering. Use it for small, fixed snapshots; future repeatable EML imports should go through SQLite. It does not implement the PST reader's spam/folder filters. Do not index its output until the command completes successfully.

## Known limitations

- The first-pass PST importer prefers plain text with HTML-to-text fallback. The separate recovery pass can try RTF; unreadable or absent representations still leave empty bodies.
- Transport headers can omit sender/recipient addresses, particularly on sent or unsent Outlook items. The recovery pass reads MAPI fields and recipient tables; unresolved addresses remain unresolved.
- Outlook conversation indexes and complete References chains can group replies; incomplete or mixed metadata can split threads. Subjects alone do not prove a shared conversation.
- Inline images can be reported as attachment filenames. Excluded documents and images cannot contribute their contents to search.
- Calendar, contact, and task objects can occur in the original import. Recovery classifies them for export filtering, but this project does not reconstruct their complete native schemas. Embedded-message contents and attachment text remain excluded.
- Signatures, quoted history, and security banners remain in extracted bodies. No semantic rewriting or automatic reply pruning is performed.
- Tests are synthetic. Native reader compatibility and extraction completeness must be checked on the actual archive, including error logs and representative samples.
- Existing v3 databases remain readable. Existing rows, fingerprints, and session IDs are not rewritten. Use a new database for uniform application of the repository edition's deduplication and threading rules. An older populated database receives a warning when its first PST snapshot baseline is recorded.

See [Funes lifecycle behavior](funes-workflow.md) before re-indexing edited data or publishing it.
