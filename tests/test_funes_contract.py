"""Opt-in contract test. Set EMAIL_FUNES_TEST_BIN to an absolute Funes binary."""
from __future__ import annotations

import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]


@unittest.skipUnless(os.environ.get("EMAIL_FUNES_TEST_BIN"), "requires explicit Funes test binary")
class FunesContractTest(unittest.TestCase):
    def test_enriched_export_and_optional_recall(self):
        binary = Path(os.environ["EMAIL_FUNES_TEST_BIN"])
        self.assertTrue(binary.is_absolute() and binary.is_file())
        with tempfile.TemporaryDirectory() as tmp:
            work = Path(tmp)
            home = work / "home"
            home.mkdir(mode=0o700)
            env = dict(os.environ, HOME=str(home), USERPROFILE=str(home),
                       FUNES_HOME=str(home / "email"), HF_HOME=str(home / "hf"),
                       XDG_CONFIG_HOME=str(home / "config"), CODEX_HOME=str(home / "codex"),
                       HF_HUB_DISABLE_IMPLICIT_TOKEN="1", FUNES_NO_TELEMETRY="1")
            # Explicitly allow a test model cache to avoid repeated downloads.
            env.pop("HF_HUB_CACHE", None)
            if os.environ.get("EMAIL_FUNES_TEST_MODEL_CACHE"):
                env["HF_HUB_CACHE"] = os.environ["EMAIL_FUNES_TEST_MODEL_CACHE"]

            def run(*args, timeout=90):
                result = subprocess.run(args, cwd=work, env=env, stdout=subprocess.PIPE,
                                        stderr=subprocess.STDOUT, encoding="utf-8", timeout=timeout)
                self.assertEqual(result.returncode, 0, result.stdout)
                return result.stdout

            db = work / "demo.sqlite"
            output = work / "export"
            run(sys.executable, ROOT / "examples/create_demo_db.py", "--enriched", "--db", db)
            run(sys.executable, ROOT / "scripts/sqlite_to_funes.py", "--db", db, "--out", output)
            checked = run(binary, "index", "--check", output)
            self.assertIn("3 turns, 3 chunks", checked)
            self.assertIn("0 rejected", checked)
            if os.environ.get("EMAIL_FUNES_TEST_RECALL") == "1":
                run(binary, "index", "--multilingual", output, timeout=600)
                recalled = run(binary, "recall", "What did we decide about the inspection?",
                               "--memory", "local", "--harness", "outlook-pst", timeout=180)
                self.assertIn("coupling", recalled.lower())
                self.assertIn("outlook-pst", recalled)
