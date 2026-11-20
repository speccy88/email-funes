# Synthetic examples

The examples contain invented project correspondence and reserved `example.com` addresses. They require only Python and SQLite; no Outlook, PST file, native parser, or account is needed.

## Try the enriched SQLite workflow

Create a new demo database and preview its export:

```powershell
python examples/create_demo_db.py --db data/demo_enriched.sqlite --enriched
python scripts/sqlite_to_funes.py --db data/demo_enriched.sqlite --dry-run
```

This database contains:

| Content | Count |
| --- | --- |
| Email messages | 3 |
| Meeting messages | 1 |
| Calendar appointments | 1 |
| Unique retained records | 5 |
| Source occurrences, including one duplicate | 6 |

It includes synthetic recovery classification and structured To/Cc/Bcc recipients. The generator populates these directly; it does not simulate a real PST recovery run. Two attachment filenames are stored, with no attachment payloads.

Before the first export, compare the choices without changing the database:

```powershell
python scripts/sqlite_to_funes.py --db data/demo_enriched.sqlite --dry-run --include-meetings
python scripts/sqlite_to_funes.py --db data/demo_enriched.sqlite --dry-run --include-all --include-bcc
```

The default selects 3 email turns, `--include-meetings` selects 4 turns, and `--include-all` selects all 5 records. Bcc remains omitted unless requested.

Create the default email-only snapshot:

```powershell
python scripts/sqlite_to_funes.py --db data/demo_enriched.sqlite --out data/demo-enriched-funes
```

The first export stores its policy and turn mappings. To try a different policy after exporting, generate another demo database with a different filename and export from that copy. `--enriched` requires a new database path.

With Funes installed, validate the completed files:

```bash
funes index --check data/demo-enriched-funes
```

The default enriched example was checked with Funes 1.12.0 on 2026-11-20: 3 turns, 3 chunks, no rejected files, and no duplicate chunk IDs. This validates the synthetic JSONL format; it does not test real PST extraction or Hugging Face publication.

See the [Funes guide](../docs/funes-workflow.md) to index the example locally and retrieve it.

## Try the basic import and deduplication example

The earlier demo remains available without `--enriched`:

```powershell
python examples/create_demo_db.py --db data/demo.sqlite
python scripts/pst_to_sqlite_direct.py --db data/demo.sqlite --stats
python scripts/pst_to_sqlite_direct.py --db data/demo.sqlite --search "coupling"
python scripts/sqlite_to_funes.py --db data/demo.sqlite --out data/demo-funes
```

It creates 3 messages, 4 source occurrences, and 2 conversations. One email is repeated in a second folder to exercise deduplication. Running this basic generator again skips the same source occurrences.

This schema has no recovery classification, so the exporter reports its legacy all-record selection. Keep demo data in a separate database from your actual archive.

## Try the EML snapshot converter

The two files in [`eml/`](eml/) form one synthetic thread. The experimental converter reads EML messages directly, without SQLite:

```powershell
python scripts/eml_to_funes.py examples/eml data/demo-eml-funes
```

Use a new or empty output directory. This shortcut illustrates another source format; repeatable future EML imports are intended to use the shared SQLite workflow.
