# Auditing and recovering PST metadata

The initial PST import creates the message database. The diagnostic and recovery commands inspect Outlook's internal MAPI properties when transport headers leave useful fields empty. Recovery works on a new SQLite copy, so the original database remains available for comparison.

Use this stage before the first Funes export and index. Enrichment can change the text a turn would contain, while Funes treats existing turn identities as append-only.

## What recovery adds

| Information | Recovery behavior |
| --- | --- |
| Outlook message class | Stores the raw class and an item category |
| Actual sender | Tries sender SMTP properties, sender address, and available email headers |
| Represented sender | Stores the on-behalf-of identity separately |
| To / Cc / Bcc | Reads the structured recipient table, retaining type, display name, raw address, and SMTP address |
| Exchange internal addresses | Resolves only unambiguous address pairs observed in the same recovery run |
| Subject and Internet Message-ID | Fills missing values from MAPI properties |
| Body | Fills a missing email/meeting body from plain text, HTML, raw MAPI body properties, or optional RTF |
| Provenance | Records how recovered fields were obtained |
| Search | Rebuilds the FTS5 index after enrichment |

Recovery preserves nonempty original fields, message IDs, fingerprints, timestamps, attachment filenames, and existing conversation assignments. A recovered thread candidate is stored for later analysis; it does not rethread the database. Attachment payloads are never read by this workflow.

The SQLite schema documents the [recovery tables and columns](sqlite-schema.md#recovery-tables-and-provenance).

## Source requirements

Use the **same unchanged PST and the same path used for the original import**. A source occurrence is matched using the normalized PST path, folder path, and position inside the folder. Moving the PST, exporting it again, or compacting it can break that correspondence.

The scripts can run wherever their dependencies are available, but an archive first imported at a Windows path should normally be recovered on that Windows machine. Copying the file to a Linux path changes the source key. You can move the resulting enriched SQLite database to Linux for preview, JSONL export, and Funes indexing; those stages no longer need the PST.

The original importer accepts PST timestamps in the existing 1970–2100 range. Values outside it stay missing, and a message lacking both sent and received dates stays in SQLite.

File size and modification time checks detect ordinary source changes. They are not a cryptographic proof that the archive is identical. Older databases may lack an original `import_sources` baseline; read the warning and retain the original archive.

A recovery run takes **one** direct-imported PST and copies the whole input database. If the database combines multiple PSTs, records unique to other PSTs remain unclassified in this copy. The email-only export omits those records. Review matched and selected counts. Multi-pass enrichment into the same output is unsupported.

Stop any program that is writing to the PST or SQLite before starting. Recovery uses SQLite's backup API to copy the source database, including a consistent database view.

## 1. Install the dependencies

From the repository root:

```powershell
python -m pip install --only-binary=:all: -r requirements/pst.txt
python -c "import pypff; print(pypff.get_version())"
```

For RTF body recovery, install the recovery requirements as well:

```powershell
python -m pip install --only-binary=:all: -r requirements/recovery.txt
```

The import name is `pypff`; the package name is `libpff-python`. The optional RTF decoder is `striprtf`. Use [installation troubleshooting](troubleshooting.md) if your Python/platform combination has no compatible wheel.

## 2. Audit a sample or the full PST

A sample checks up to 40 distributed items in each folder:

```powershell
python scripts/pst_metadata_audit.py --pst "C:\MailArchives\archive.pst" --db data/email.sqlite --out data/pst_metadata_audit.log
```

For a complete audit:

```powershell
python scripts/pst_metadata_audit.py --pst "C:\MailArchives\archive.pst" --db data/email.sqlite --out data/pst_metadata_audit_full.log --full
```

The database is optional. Supplying it lets the audit compare stored fields against the native properties of matching PST items.

| Option | Meaning |
| --- | --- |
| `--pst PATH` | Existing PST, opened for reading |
| `--db PATH` | Optional existing SQLite, opened read-only |
| `--out PATH` | Aggregate diagnostic log |
| `--per-folder N` | Distributed samples per folder; default `40`, `0` means all |
| `--full` | Inspect every item |
| `--progress N` | Progress interval; default `2000` inspected items |
| `--max-report-folders N` | Maximum anonymized folder rows; default `20` |

The report describes property availability, item categories, recipient structures, read errors, and matches to existing source records. Folder labels are anonymized for each run. It does not include message subjects, bodies, names, or addresses.

A populated MAPI property means a candidate is available; it does not guarantee that its value is usable, is an SMTP address, or differs from another property. Missing-field rates also depend on item category: a contact or calendar entry should not be assessed as if it were an email.

## 3. Try recovery on a small sample

The recovery command always starts from the original imported database and writes a fresh output file:

```powershell
python scripts/pst_sqlite_recover.py --pst "C:\MailArchives\archive.pst" --db data/email.sqlite --out-db data/email_enriched_test.sqlite --log data/pst_recovery_test.log --limit-matched 200 --recover-rtf
```

The test copy still contains **all original message rows**. The limit controls how many unique matching messages receive enrichment, not how many rows are copied. Unprocessed rows retain a blank category and are excluded by the default enriched JSONL selection. Use the full recovery output for the production export.

For the full recovery, start again from the original database with a different output name:

```powershell
python scripts/pst_sqlite_recover.py --pst "C:\MailArchives\archive.pst" --db data/email.sqlite --out-db data/email_enriched.sqlite --log data/pst_recovery.log --recover-rtf
```

| Option | Meaning |
| --- | --- |
| `--pst PATH` | Original unchanged PST at its original import path |
| `--db PATH` | Source SQLite; never modified |
| `--out-db PATH` | New enriched SQLite; must not exist |
| `--log PATH` | Aggregate recovery report |
| `--limit-matched N` | Enrich up to N unique matched messages; default `0` means all |
| `--max-recipients N` | Per-item recipient cap; default `4096` |
| `--progress N` | Progress interval; default `4000` inspected items |
| `--recover-rtf` | Attempt RTF body decoding using `striprtf` |

An input with existing recovery tables is rejected. To rerun or change recovery options, use the original database and a new output filename. This recovery command does not resume a partial enrichment file in place.

## 4. Read the recovery report

The report separates scanned source items, matched occurrences, unique enriched rows, and recovered recipient rows. These count different things. One retained email can have multiple source occurrences; an excluded Junk or Deleted Items entry can exist in the PST without a corresponding SQLite row.

Field-fill totals are counted only after the message's database changes succeed. A failed message rolls back its updates, recipient rows, and alias evidence. The report also records field-reading problems and text repairs without logging the underlying content. Valid UTF-16 surrogate pairs are combined; isolated surrogate code units are replaced with the Unicode replacement character (`U+FFFD`) so they can be stored safely.

When an Exchange internal address maps to several SMTP addresses, recovery leaves it unresolved. It does not guess from a display name or query a corporate directory. `recovery_aliases.evidence_count` records observations supporting a retained unambiguous mapping.

A readable body can still be absent after recovery. The source may be an attachment-only message, a non-email object, or a representation the reader cannot decode. The log records these limitations; it does not prove why a particular body is empty.

Review the log even when an enriched database is produced. Some individual read errors permit the run to finish with usable partial recovery; structural or source-change failures prevent publication of the completed output.

| Exit code | Meaning |
| --- | --- |
| `0` | Completed without recorded issues |
| `2` | Completed and produced an enriched database, with issues or text repairs to review |
| `1` | Aborted; no enriched output published |
| `130` | Interrupted; no enriched output published |

Invalid command-line arguments also return `2` before recovery starts; check whether the completion message and output file exist. Failed runs write a sanitized failure log when the destination is writable.

The output is finalized atomically without replacing another file. Its filesystem must support hard links, such as NTFS or a usual Linux filesystem. If final publication fails on another filesystem, use a local supported volume and rerun from the original database.

## 5. Review and prepare JSONL

Use the enriched database for the next stage:

```powershell
python scripts/pst_to_sqlite_direct.py --db data/email_enriched.sqlite --stats
python scripts/sqlite_to_funes.py --db data/email_enriched.sqlite --dry-run
```

The dry run is read-only and reports selection and date-availability counts without generating JSONL or assigning turn identities. It checks policy compatibility but does not fully validate content or timestamp syntax. The default selects the recovered `Email / IPM.Note` category; `--include-meetings` also selects `Meeting request`. This category includes several meeting-message subclasses, not just invitations.

Decide whether to include meetings and Bcc before the first export. The exporter stores its selection/rendering policy with `funes_map`, uses normalized recipients where available, and rejects incompatible later choices. See the [converter reference](converters.md#sqlite-to-funes).

## Limits and recovery decisions

- Recovery enriches records already retained by the importer. It does not add previously excluded emails or recompute duplicate identities.
- Native recipient tables can still be absent, capped, or partially unreadable. Some recipient entries contain names or internal addresses without an SMTP value.
- The parser reads stored archive data. It does not contact Outlook, Exchange, or an address directory to resolve missing fields.
- RTF conversion extracts text; it does not reproduce document layout or attachment content.
- Saved fingerprints remain the original import fingerprints even when blank fields are filled. Do not treat the enriched copy as a fresh import under different deduplication rules.
- Tests use synthetic MAPI objects. Inspect real extraction results and aggregate errors before treating the archive as complete.
- If the original database already contains Funes mappings, recovery preserves them. Updated text under those IDs will not update an existing Funes index; plan a new derived export/index deliberately.
