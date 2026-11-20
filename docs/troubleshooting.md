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

Some Outlook items have no transport headers. Their sender address, To/Cc, or Message-ID can be unavailable even when the subject/body exists. RTF-only messages can lack a usable body. The importer reports missing fields; the exporter skips records with no sent or received timestamp.

Check the corresponding messages in the original archive. A missing field is an extraction limitation, not evidence that the original message lacked useful information.

## “Output must be a new or empty directory”

Use a new export directory:

```powershell
python scripts/sqlite_to_funes.py --db data/email.sqlite --out data/funes-export-02
```

This avoids stale JSONL files from an earlier, larger export. Keep `funes_map` in the same database so turn identities remain stable between snapshots.

## Funes cannot answer from an attachment

Only its filename was stored. The email body may mention the document, but the attachment contents were intentionally excluded. Import selected documents separately if that becomes a requirement.

## Edited email text does not change Funes results

Funes deduplicates by turn/block identity rather than text. An edit under an existing turn ID is not an update. Use explicit new versions or rebuild the derived memory. Deleting an email from SQLite does not remove indexed or published Funes history. See the [official identity rules](https://github.com/huggingface/funes/blob/main/docs/funes-jsonl.md#identity--what-you-owe-the-store).

## Moving an existing v3 database

The existing message/source tables are compatible. The repository edition adds the source-snapshot baseline and more conservative duplicate/threading rules. It does not rewrite older rows. For uniform behavior, import into a fresh database; to continue an old sample, retain the original unchanged source and review the baseline warning.

If you already indexed the earlier JSONL, keep that memory separate while validating a fresh export. Never assume a new conversion automatically replaces an old Funes memory.
