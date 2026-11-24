# Converter reference

Install the CLI from the source checkout. The original scripts remain available for compatibility. [Installation and compatibility](../README.md#install-and-compatibility)

## Direct PST reader

```powershell
email-funes import "C:\MailArchives\archive.pst" --format pst --db data/email.sqlite --limit 100
```

The reader uses `pypff`, provided by the official [`libpff-python`](https://pypi.org/project/libpff-python/) package. It opens the PST directly. It does not call Outlook, send mail, or upload content.

| Option | Default | Meaning |
| --- | --- | --- |
| `ARCHIVE` | Required for import | Existing PST archive (the original script uses `--pst PATH`) |
| `--db PATH` | Required | SQLite destination |
| `--limit N` | `0` | Maximum inspected items; `0` means all |
| `--batch-size N` | `250` | Commit completed records every N inspected items |
| `--progress N` | `1000` | Print progress every N inspected items |
| `--include-junk` | Off | Include recognized Junk/Spam folders |
| `--include-deleted` | Off | Include recognized Deleted Items/Trash folders |
| `--include-header-spam` | Off | Keep messages marked by explicit spam rules |
| `--skip-bulk` | Off | Exclude List-Id, List-Unsubscribe, and bulk/list Precedence messages |
| `--max-body-chars N` | `0` | Optional storage truncation; `0` retains the full extracted body |
| `--private-diagnostics` | Off | Include private location and exception details in the local import error log |

Use `email-funes stats --db data/email.sqlite` for read-only counts. The compatibility script also accepts `--stats`, `--search QUERY`, and `--results N` (default 15). Search displays matching email content.

Folder names are normalized for English and French names. Folder filtering also excludes their descendants. Deleted Items is a separate choice from spam; include it if deleted correspondence belongs in your project history.

The spam rules recognize `X-Spam-Flag: YES`, `X-Spam-Status: Yes`, Exchange SCL values of 5 or more, and subjects beginning `[SPAM]` or `[JUNK]`. They do not independently determine whether a message is malicious. `--skip-bulk` can remove useful technical mailing lists.

### Duplicates and resume behavior

Duplicates are matching **retained fields**: message ID, sender metadata, recipients, timestamp, subject, full extracted body, attachment names, and attachment count. Different recipients, message IDs, or body variants are preserved. Attachment bytes are not compared. The full body is fingerprinted before optional truncation.

Each source occurrence is recorded separately in `source_items`, so one email can retain several folder references. Known source entries are skipped on restart.

The direct reader's resume identifiers use the PST path, folder path, and item position. `import_sources` records file size and modification time to detect ordinary changes to that snapshot. If the PST changes, import the new snapshot into a new database. This is a metadata guard, not a full-file cryptographic hash. Do not edit a PST during extraction.

Filtered items are counted but not stored as email rows or durable exclusion records. They may be checked again on a later run. This is not continuous mailbox synchronization.

### Error handling

On an extraction error the direct reader writes a local log beside the database, using `*.direct.errors.txt` or a fresh suffixed filename if a previous log exists. Default diagnostics contain event labels, counts and exception types. Add `--private-diagnostics` for a private error log with location and exception details. Keep these private logs out of public issue reports. The importer continues past individual unreadable records and returns a nonzero status on extraction errors. Ctrl+C commits completed records and returns 130.

For a clean import using different filter or truncation settings, create a new database. Existing rows are not rewritten by changing command-line flags.

## Outlook Classic alternative

```powershell
email-funes import "C:\MailArchives\archive.pst" --format pst --backend outlook --db data/outlook-com.sqlite --limit 100
```

This Windows-only adapter uses Outlook Classic's COM object model. It needs a usable Outlook profile. The new Outlook application does not supply this interface. Open Outlook Classic and the PST manually if the adapter reports that it is not connected.

This adapter is retained for compatibility with the initial scripts. It may temporarily mount a PST in the Outlook profile; Outlook can update local PST metadata. The direct reader is the main workflow.

## PST metadata audit and recovery

Use `pst_metadata_audit.py` to measure native property coverage and `pst_sqlite_recover.py` to enrich a fresh copy of an existing import. These commands share the `pypff` dependency. Recovery can also decode RTF text with optional `striprtf`.

```powershell
email-funes audit --pst "C:\MailArchives\archive.pst" --db data/email.sqlite --out data/audit.log --full
email-funes recover --pst "C:\MailArchives\archive.pst" --db data/email.sqlite --out-db data/email_enriched.sqlite --log data/recovery.log --recover-rtf
```

The original PST path and immutable snapshot are part of source matching. Recovery does not add filtered source items or rewrite existing message identities. It adds item classification, structured recipients, recovered fields, and provenance to the copy. Its aggregate log contains no message text or addresses.

Choose fresh paths for both reports and the enriched database. Audit reports sample counts and labels folders with opaque identifiers. Recovery uses the native direct reader's occurrence keys. Outlook COM imports use a different occurrence identity.

Several immutable PSTs may feed one first-pass SQLite database. Recovery currently accepts **one PST per fresh copy**. In a database containing several PSTs, records exclusive to other archives stay unclassified and are excluded by email-only selection. An enriched copy cannot be used as the input for a second recovery pass. For full metadata recovery across independent archives, recover each in its own database and use separate Funes memories if conversations overlap.

See the [metadata recovery guide](metadata-recovery.md) for dependencies, flags, pilot runs, source checks, and limitations.

## SQLite to Funes

```powershell
email-funes export --db data/email_enriched.sqlite --dry-run
email-funes export --db data/email_enriched.sqlite --out data/funes-export
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

The first preview or export of an unclassified database requires a selection decision. Run recovery for email-only selection, or pass `--include-all` to count or export every retained Outlook item, including contacts and calendar entries. A compatible earlier all-record policy or populated legacy mapping supplies that decision on later runs.

Normalized recipient rows take precedence over the original flattened To/Cc text. If a recipient type has no normalized rows, the exporter uses the original field where available. Existing but empty typed rows do not fall back to an older display list. Unresolved recipients keep their display names. Recovered Bcc is retained locally in SQLite. Export omits the Bcc field by default. An address quoted in a body or appearing elsewhere in the message remains in the text.

Only the selected records receive new Funes mappings. Records with neither a sent nor a received date stay in SQLite and are skipped; timestamps are never invented. Records with an empty body can still provide searchable subjects, participants, dates, and attachment filenames.

### Stable export state

A normal export adds two tables to SQLite:

- `funes_map` keeps each selected record's `session_id`, `turn_uuid`, and dense `seq`.
- `funes_export_policy` keeps the selection and recipient-rendering choices, including Bcc inclusion.

Back up these tables together with the messages. Later-discovered emails append to their session's sequence even when their timestamp is older. Files can contain several sessions, and a session can span files.

Omit selection flags to reuse the saved selection. A database exported with Bcc requires `--include-bcc` again on every preview or export. A conflicting selection, Bcc choice, or incompatible schema fails before publication. A populated older mapping without a policy can continue under its original unclassified, all-record rendering. Enriching an already mapped legacy database changes the rendering inputs and requires a deliberate rebuild.

If a policy change is necessary, preserve the existing database/export and start from a separate copy of the archive taken **before its first mapping**. Use a separate derived Funes memory for a rebuild. Do not delete mapping rows in a published workflow to force the command through: that can shift sequences or retain old indexed content under reused identities.

### Output and validation

Choose a new or empty directory in a private parent directory. The exporter stages and validates the complete snapshot, commits its mappings, then publishes the directory with a rename. Rendering errors roll back new assignments and remove staging files. If the final rename fails after the commit, the database retains valid assignments for a retry. A new snapshot folder can reuse those IDs.

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
