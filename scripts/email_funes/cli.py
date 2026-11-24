"""The installed email-funes command."""

import argparse
from contextlib import closing
from pathlib import Path
import sqlite3
import sys

from email_funes_files import FileSafetyError
from .adapters import ARCHIVE_ADAPTERS, pst_options


def import_archive(argv, *, prog):
    parser = argparse.ArgumentParser(prog=prog, description="Import a local archive into the shared SQLite database.")
    parser.add_argument("archive", metavar="ARCHIVE", type=Path, help="Existing local Outlook PST file")
    parser.add_argument("--format", choices=ARCHIVE_ADAPTERS, default="pst", help="Archive format (default pst)")
    parser.add_argument("--backend", choices=("libpff", "outlook"), default="libpff",
                        help="PST reader (default libpff). Outlook requires Windows and Outlook Classic.")
    pst_options(parser)
    parser.set_defaults(stats=False, search=None, results=15)
    args = parser.parse_args(argv)
    return ARCHIVE_ADAPTERS[args.format](args, parser)


def stats(argv, *, prog):
    from email_funes_common import show_stats
    from email_funes_files import open_readonly

    parser = argparse.ArgumentParser(prog=prog, description="Read aggregate counts from an existing SQLite archive.")
    parser.add_argument("--db", required=True, type=Path, help="Existing email-funes SQLite file")
    args = parser.parse_args(argv)
    with closing(open_readonly(args.db.expanduser())) as db:
        show_stats(db)
    return 0


def main(argv=None) -> int:
    from pst_metadata_audit import main as audit
    from pst_sqlite_recover import main as recover
    from sqlite_to_funes import main as export

    commands = {
        "import": (import_archive, "Read an Outlook PST into SQLite"),
        "audit": (audit, "Inspect PST metadata and write an aggregate report"),
        "recover": (recover, "Recover metadata into a new SQLite copy"),
        "export": (export, "Preview selection or write Funes JSONL"),
        "stats": (stats, "Read SQLite message and source counts"),
    }
    parser = argparse.ArgumentParser(prog="email-funes", description="Convert local email archives to Funes turns.")
    subcommands = parser.add_subparsers(dest="command", required=True, metavar="COMMAND")
    for name, (_, description) in commands.items():
        subcommands.add_parser(name, help=description)
    args = sys.argv[1:] if argv is None else argv
    command = parser.parse_args(args[:1]).command
    run, _ = commands[command]
    try:
        return run(args[1:], prog=f"email-funes {command}")
    except KeyboardInterrupt:
        return 130
    except FileSafetyError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except (OSError, sqlite3.Error) as exc:
        print(f"{command} failed ({type(exc).__name__}). Check the input, permissions and available space.",
              file=sys.stderr)
        return 1
