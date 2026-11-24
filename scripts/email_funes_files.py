"""Local files used by the converter. Never change the process-wide umask."""
from __future__ import annotations

import os
from contextlib import contextmanager, ExitStack
from pathlib import Path
import shutil
import sqlite3
import stat
import tempfile
import unicodedata


class FileSafetyError(ValueError):
    """Only fixed, public diagnostic text belongs in these exceptions."""


def private_mkdir(path: Path) -> None:
    """Apply the private mode to each directory we create, leaving others alone."""
    if path.is_dir():
        return
    if path.parent != path:
        private_mkdir(path.parent)
    path.mkdir(mode=0o700, exist_ok=True)


def new_private_text(path: Path):
    private_mkdir(path.parent)
    return open(path, "x", encoding="utf-8", newline="\n", opener=_private_open)


def check_destination(path: Path) -> None:
    try:
        info = path.lstat()
    except FileNotFoundError:
        return
    reparse = getattr(info, "st_file_attributes", 0) & stat.FILE_ATTRIBUTE_REPARSE_POINT
    if reparse or not stat.S_ISDIR(info.st_mode) or any(path.iterdir()):
        raise FileSafetyError("Output must be a new or empty directory without links.")


@contextmanager
def staged_directory(out_dir: Path, *, before_publish=None):
    """Publish all output together on the destination filesystem."""
    check_destination(out_dir)
    private_mkdir(out_dir.parent)
    staging = Path(tempfile.mkdtemp(prefix=".email-funes-export-", dir=out_dir.parent))
    try:
        yield staging
        check_destination(out_dir)
        if before_publish is not None:
            before_publish()
        if os.name == "nt" and out_dir.exists():
            # Windows cannot rename a directory over an existing empty one.
            out_dir.rmdir()
        staging.rename(out_dir)
    finally:
        if staging.exists():
            shutil.rmtree(staging)


def _private_open(name, flags):
    return os.open(name, flags | getattr(os, "O_NOFOLLOW", 0), 0o600)


def sqlite_paths(path: Path) -> tuple[Path, ...]:
    return (path, *(Path(str(path) + suffix) for suffix in ("-wal", "-shm", "-journal")))


def require_distinct(*paths: Path) -> None:
    names = set()
    files = set()
    for path in paths:
        name = path.resolve()
        if any(name == existing or name.is_relative_to(existing) or existing.is_relative_to(name)
               for existing in names):
            raise FileSafetyError("Source, database, SQLite sidecars and output paths must be distinct.")
        names.add(name)
        if path.exists():
            info = path.stat()
            key = (info.st_dev, info.st_ino)
            if info.st_ino and key in files:
                raise FileSafetyError("Input and output files must not be aliases or hard links.")
            files.add(key)


def check_writable_file(path: Path) -> None:
    try:
        info = path.lstat()
    except FileNotFoundError:
        return
    if not stat.S_ISREG(info.st_mode) or info.st_nlink > 1:
        raise FileSafetyError("Use a regular output file with no symbolic or hard links.")


def protect_sqlite(path: Path) -> None:
    paths = sqlite_paths(path)
    for member in paths:
        check_writable_file(member)
    if not path.exists() and any(p.exists() for p in paths[1:]):
        raise FileSafetyError("Orphan SQLite sidecars found. Choose a fresh database path.")
    # SQLite creates journals and WAL/SHM with the database's permissions.
    if os.name == "posix":
        for member in paths:
            if member.exists() and stat.S_IMODE(member.stat().st_mode) & 0o077:
                member.chmod(stat.S_IMODE(member.stat().st_mode) & ~0o077)


def create_private_database(path: Path) -> None:
    protect_sqlite(path)
    private_mkdir(path.parent)
    try:
        fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
    except FileExistsError:
        check_writable_file(path)
    else:
        os.close(fd)


@contextmanager
def open_import_log(path: Path):
    """Write one fresh diagnostic file only if this run encounters an error."""
    check_writable_file(path)
    with ExitStack() as stack:
        stream = None

        def write(line):
            nonlocal stream
            if stream is None:
                try:
                    handle = new_private_text(path)
                except FileExistsError:
                    # Keep previous logs, including private reports from older versions.
                    handle = tempfile.NamedTemporaryFile(
                        mode="w", encoding="utf-8", newline="\n", delete=False,
                        prefix=path.stem + "-", suffix=".txt", dir=path.parent)
                stream = stack.enter_context(handle)
            print(line, file=stream)

        yield write


def _snapshot_state(path: Path) -> tuple:
    return (*source_signature(path),
            *(os.path.lexists(p) for p in sqlite_paths(path)[1:]))


def source_signature(path: Path) -> tuple[int, int, int, int]:
    info = path.stat()
    return info.st_size, info.st_mtime_ns, info.st_dev, info.st_ino


class _ClosedArchive(sqlite3.Connection):
    def close(self):
        super().close()
        if self.initial_state is None:
            return
        before, self.initial_state = self.initial_state, None
        try:
            after = _snapshot_state(self.archive)
        except OSError:
            after = None
        if after != before:
            raise FileSafetyError("The database changed during reading. Close writers and retry.")


def open_readonly(path: Path) -> sqlite3.Connection:
    """Read a closed archive without creating sidecars, or use SQLite's live WAL locks."""
    path = path.resolve(strict=True)
    _, wal, shm, journal = sqlite_paths(path)
    for sidecar in (wal, shm, journal):
        check_writable_file(sidecar)
    state = _snapshot_state(path)
    if journal.exists():
        with journal.open("rb") as stream:
            if any(stream.read(28)):
                raise FileSafetyError("An active rollback journal exists. Close the writer and retry.")
    uri = path.as_uri() + "?mode=ro"
    if wal.exists() != shm.exists():
        raise FileSafetyError("Incomplete SQLite WAL sidecars. Close the writer or restore the complete archive.")
    # A normal read-only open of a *closed* WAL archive can create empty WAL
    # and SHM files. Read an immutable snapshot only while it remains quiescent.
    with path.open("rb") as stream:
        wal_mode = stream.read(20)[18:20] == b"\x02\x02"
    if wal_mode and not wal.exists():
        db = sqlite3.connect(uri + "&immutable=1", uri=True, factory=_ClosedArchive)
        db.archive = path
        db.initial_state = state
    else:
        db = sqlite3.connect(uri, uri=True)
    db.execute("PRAGMA query_only=ON")
    return db


def display_text(value) -> str:
    """Escape terminal control and bidirectional formatting characters at display time."""
    return "".join(ascii(c)[1:-1] if unicodedata.category(c) in ("Cc", "Cf", "Cs", "Zl", "Zp")
                   else c for c in str(value))


def diagnostic(event, exc, *, private=False, location="") -> str:
    line = f"{event} error={type(exc).__name__}"
    if private:
        line += f" location={display_text(location)} detail={display_text(exc)}"
    return line
