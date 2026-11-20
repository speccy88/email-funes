# Troubleshooting

## pip cannot find the PST reader

Use the official package name `libpff-python`, which imports as `pypff`:

```powershell
python --version
python -m pip --version
python -m pip install --only-binary=:all: -r requirements/pst.txt
python -c "import pypff; print(pypff.get_version())"
```

The [PyPI release files](https://pypi.org/project/libpff-python/#files) list supported interpreter/platform combinations. If no wheel matches, inspect the Python version and architecture:

```powershell
python -c "import sys, platform; print(sys.executable); print(sys.version); print(platform.machine())"
```

A managed package index can also lack the package. Check that locally with `python -m pip config list`; its output can contain private repository URLs or credentials, so redact it before sharing. Use the approved package source for your environment.

Do not use the earlier `libpff-python-windows` recommendation. This repository uses the official package.

## Outlook says “Vous n'êtes pas connecté”

This concerns the optional COM importer, not SQLite. A local PST can still require a working Outlook session when accessed through COM. Open Outlook Classic with its configured profile and open the PST before retrying, or use `pst_to_sqlite_direct.py` to read the file directly.

The direct reader requires no Outlook session. Close Outlook if it has the PST locked.

## The database contains only 10 or 100 emails

Check whether the command used `--limit`. Remove it to process all items from the same unchanged archive. Already imported occurrences count toward the inspection limit and are skipped.

```powershell
python scripts/pst_to_sqlite_direct.py --db data/email.sqlite --stats
```

Folder and spam filtering also affect counts. Inspect the run summary and error log. A small sample does not establish full-archive completeness.

## Resume refuses a changed PST

The direct reader tracks the source file's size and modification time because its resume keys contain item positions. Use the unchanged original for resume. A new export, compacted PST, or modified archive should go into a new SQLite database. Do not bypass the guard by editing its stored baseline.

## Missing addresses, dates, or bodies

Some Outlook items have incomplete transport headers but retain metadata in MAPI properties or a native recipient table. Run the [metadata audit and recovery](metadata-recovery.md) against the unchanged original PST. It fills missing fields in a new SQLite copy and reports unresolved cases.

Optional `--recover-rtf` needs the packages in `requirements/recovery.txt`. RTF extraction can recover additional text, but a remaining empty body is not proof of a damaged message: attachment-only emails and non-email objects can also have empty bodies. Check the relevant item in the source archive if its contents matter.

Recovery does not invent dates. The exporter skips records with neither a sent nor a received timestamp.

## Recovery matches no source items

Check that the PST is the original unchanged snapshot at the path used for the first import. Source keys include the PST path, folder path, and item position. A Windows-to-Linux path change will generally prevent matching even if the file's bytes are identical.

Run the recovery where the original path is available, then move the enriched SQLite to Linux for JSONL export and Funes. Do not bypass a snapshot-change error by editing `source_items` or the stored baseline.

## Recovery test output looks as large as the full database

`--limit-matched` limits enrichment, not the initial SQLite backup. A pilot copy contains all original records but only a subset has recovery metadata. Use the report's unique-enriched count or count rows in `recovery_metadata` to distinguish a pilot from a full recovery.

For the full run, pass the original un-enriched database and a new output name. An existing enriched input or output is rejected; recovery does not resume a pilot copy in place. Use the full recovery output for production JSONL: unprocessed pilot rows have no category and are excluded by the default email selection.

## Export policy differs from the previous run

The first export stores its selection and recipient-rendering policy alongside `funes_map`. Later exports must use the same choices, including `--include-meetings`, `--include-all`, and `--include-bcc`. A new output folder does not reset the database's policy.

For a deliberate policy change, keep the existing database and derive a separate export from a copy made before its first mapping. Build a separate Funes memory when the changed policy affects already-indexed text. Do not remove mapped rows or renumber sequences to bypass the check.

An enriched database carrying an earlier legacy mapping also needs this deliberate rebuild decision. Enrichment can change text under existing IDs, which Funes will not replace automatically.

## “Output must be a new or empty directory”

Use a new export directory:

```powershell
python scripts/sqlite_to_funes.py --db data/email_enriched.sqlite --out data/funes-export-02
```

This avoids stale JSONL files from an earlier, larger export. Keep `funes_map` in the same database so turn identities remain stable between snapshots.

## Funes cannot answer from an attachment

Only its filename was stored. The email body may mention the document, but the attachment contents were intentionally excluded. Import selected documents separately if that becomes a requirement.

## Edited email text does not change Funes results

Funes deduplicates by turn/block identity rather than text. An edit under an existing turn ID is not an update. Use explicit new versions or rebuild the derived memory. Deleting an email from SQLite does not remove indexed or published Funes history. See the [official identity rules](https://github.com/huggingface/funes/blob/main/docs/funes-jsonl.md#identity--what-you-owe-the-store).

## Moving an existing v3 database

The existing message/source tables are compatible. The repository edition adds the source-snapshot baseline and more conservative duplicate/threading rules. It does not rewrite older rows. For uniform behavior, import into a fresh database; to continue an old sample, retain the original unchanged source and review the baseline warning.

If you already indexed the earlier JSONL, keep that memory separate while validating a fresh export. Never assume a new conversion automatically replaces an old Funes memory.
