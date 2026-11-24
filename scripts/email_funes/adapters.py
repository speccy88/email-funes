"""Local archive adapters. Each reader writes the shared SQLite message model."""

import argparse
from pathlib import Path


PST_INSTALL_HELP = (
    'The native PST reader (libpff-python, imported as pypff) is unavailable. '
    'From the email-funes checkout, run python -m pip install --only-binary=:all: ".[pst]" '
    'using a Python version and platform with a matching wheel.'
)


def pst_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--db", required=True, type=Path, help="Destination SQLite file")
    parser.add_argument("--limit", type=int, default=0, help="Inspect at most N items (0 for all)")
    parser.add_argument("--batch-size", type=int, default=250, help="Commit every N items (default 250)")
    parser.add_argument("--progress", type=int, default=1000, help="Report progress every N items (default 1000)")
    parser.add_argument("--include-junk", dest="exclude_junk", action="store_false", help="Read Junk/Spam folders")
    parser.add_argument("--include-deleted", dest="exclude_deleted", action="store_false", help="Read Deleted Items folders")
    parser.add_argument("--include-header-spam", action="store_true", help="Keep items marked as spam by headers")
    parser.add_argument("--skip-bulk", action="store_true", help="Omit newsletters and mailing lists")
    parser.add_argument("--max-body-chars", type=int, default=0, help="Optional body length limit (0 for unlimited)")
    parser.add_argument("--private-diagnostics", action="store_true",
                        help="Write detailed private diagnostics locally. Keep these logs out of public issue reports.")


def legacy_pst_options(parser: argparse.ArgumentParser) -> None:
    pst_options(parser)
    parser.add_argument("--pst", type=Path, help="Existing local PST file")
    parser.add_argument("--stats", action="store_true", help="Show statistics for an existing database")
    parser.add_argument("--search", metavar="FTS5_QUERY", help="Search an existing database")
    parser.add_argument("--results", type=int, default=15, help="Maximum search hits (default 15)")


def import_pst(args: argparse.Namespace, parser: argparse.ArgumentParser) -> int:
    args.pst = args.archive
    if args.backend == "outlook":
        from pst_to_sqlite import run
    else:
        from pst_to_sqlite_direct import run
    return run(args, parser)


ARCHIVE_ADAPTERS = {"pst": import_pst}
