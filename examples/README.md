# Synthetic examples

These examples contain invented correspondence and `example.com` addresses. They are safe to use for learning the workflow without opening a personal mailbox. None of these commands uploads data.

## Try SQLite and Funes JSONL without a PST

Run from the repository root with Python 3.10 or newer:

```powershell
python examples/create_demo_db.py --db data/demo.sqlite
python scripts/pst_to_sqlite_direct.py --db data/demo.sqlite --stats
python scripts/pst_to_sqlite_direct.py --db data/demo.sqlite --search "coupling"
python scripts/sqlite_to_funes.py --db data/demo.sqlite --out data/demo-funes
```

A new demo database contains **3 messages, 4 source occurrences, and 2 conversations**. One message is repeated in a second source folder to exercise deduplication. Two attachment filenames are stored; there are no attachment files or binary contents.

Running the demo generator again skips the same four source occurrences. Use a separate demo database so these records do not mix with your real archive.

The JSONL export can be validated with an installed Funes CLI:

```bash
funes index --check data/demo-funes
```

See the main [README](../README.md) for local indexing and optional Hugging Face synchronization.

## Try the earlier EML snapshot converter

The two files in [`eml/`](eml/) form one synthetic email thread. The earlier converter exports EML messages directly to Funes JSONL, without the SQLite stage:

```powershell
python scripts/eml_to_funes.py examples/eml data/demo-eml-funes
```

Use a new, empty output directory for that snapshot converter. It is retained as an example of another input format; the SQLite route is the recommended basis for repeatable imports.
