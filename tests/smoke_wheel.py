"""Exercise the installed console command offline. Usage: python tests/smoke_wheel.py dist"""

from email.parser import BytesParser
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import venv
import zipfile


ROOT = Path(__file__).resolve().parents[1]
MODULES = {
    "email_funes/__init__.py", "email_funes/__main__.py", "email_funes/adapters.py",
    "email_funes/cli.py", "email_funes_common.py", "email_funes_files.py", "eml_to_funes.py",
    "pst_metadata_audit.py", "pst_sqlite_recover.py", "pst_to_sqlite.py",
    "pst_to_sqlite_direct.py", "sqlite_to_funes.py",
}
METADATA = {"entry_points.txt", "METADATA", "RECORD", "WHEEL", "licenses/LICENSE"}


def check_wheel(wheel):
    with zipfile.ZipFile(wheel) as archive:
        files = set(archive.namelist())
        metadata = {p for p in files if ".dist-info/" in p}
        if files - metadata != MODULES or {p.split("/", 1)[1] for p in metadata} != METADATA:
            raise AssertionError("The wheel differs from the reviewed list of modules and package metadata")
        (metadata_file,) = (p for p in metadata if p.endswith("/METADATA"))
        headers = BytesParser().parsebytes(archive.read(metadata_file), headersonly=True)
        assert headers["License-Expression"] == "MIT"
        assert headers.get_all("License-File") == ["LICENSE"]
        dist_info = metadata_file.rsplit("/", 1)[0]
        assert archive.read(f"{dist_info}/licenses/LICENSE") == (ROOT / "LICENSE").read_bytes()

    with tempfile.TemporaryDirectory() as temp:
        cwd = Path(temp)
        env = os.environ.copy()
        for key in ("PYTHONHOME", "PYTHONPATH"):
            env.pop(key, None)
        env["PYTHONUTF8"] = "1"
        env["PIP_DISABLE_PIP_VERSION_CHECK"] = "1"

        def run(*args, ok=True):
            result = subprocess.run(args, env=env, cwd=cwd, capture_output=True, text=True, encoding="utf-8")
            output = result.stdout + result.stderr
            if (result.returncode == 0) != ok:
                raise AssertionError(f"Unexpected exit status {result.returncode}\n{output}")
            return output

        venv.EnvBuilder(with_pip=True).create(cwd / "venv")
        bin_dir = cwd / "venv" / ("Scripts" if os.name == "nt" else "bin")
        python = bin_dir / ("python.exe" if os.name == "nt" else "python")
        cli = bin_dir / ("email-funes.exe" if os.name == "nt" else "email-funes")
        run(python, "-I", "-m", "pip", "install", "--no-index", "--no-deps", wheel)
        run(cli, "--help")
        for command in ("import", "audit", "recover", "export", "stats"):
            assert f"email-funes {command}" in run(cli, command, "--help")
        run(python, "-I", "-m", "email_funes", "--help")
        for name in ("pst_to_sqlite_direct", "pst_to_sqlite", "pst_metadata_audit", "pst_sqlite_recover",
                     "sqlite_to_funes", "eml_to_funes"):
            run(python, "-I", "-m", name, "--help")

        # Only the fixture generator runs from the checkout. All converter
        # invocations below resolve through the wheel's console entry point.
        run(python, "-I", ROOT / "examples" / "create_demo_db.py", "--db", "demo.sqlite", "--enriched")
        fixture = cwd / "demo.sqlite"
        original = fixture.read_bytes()
        assert "Stored records" in run(cli, "stats", "--db", "demo.sqlite")
        run(cli, "export", "--db", "demo.sqlite", "--dry-run")
        assert fixture.read_bytes() == original
        run(cli, "export", "--db", "demo.sqlite", "--out", "jsonl")
        turns = [json.loads(line) for file in (cwd / "jsonl").glob("*.funes.jsonl")
                 for line in file.read_text(encoding="utf-8").splitlines()]
        assert len(turns) == 3 and all(t["harness"] == "outlook-pst" for t in turns)
        (cwd / "synthetic.pst").touch()
        for args in (
            ("import", "synthetic.pst", "--format", "pst", "--db", "import.sqlite"),
            ("audit", "--pst", "synthetic.pst", "--out", "audit.log"),
            ("recover", "--pst", "synthetic.pst", "--db", "demo.sqlite", "--out-db", "recovered.sqlite"),
        ):
            output = run(cli, *args, ok=False)
            assert "libpff-python" in output and "Traceback" not in output, output
        assert not (cwd / "import.sqlite").exists()
        if os.name == "nt":
            output = run(cli, "import", "synthetic.pst", "--backend", "outlook", "--db", "com.sqlite", ok=False)
            assert "pywin32" in output and "Traceback" not in output, output


if __name__ == "__main__":
    (artifact,) = Path(sys.argv[1]).resolve().glob("email_funes-*.whl")
    check_wheel(artifact)
    print("Base wheel installed and exercised successfully outside the checkout.")
