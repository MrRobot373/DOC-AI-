"""CLI exit-code behavior (mock LLM, no network)."""

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import cli
from test_engine_pipeline import FakeOllamaClient, _build_doc


class CliExitCodeTests(unittest.TestCase):
    def setUp(self):
        # Force the CLI to use the deterministic mock client.
        self._orig = cli._build_client
        cli._build_client = lambda api_key, host: FakeOllamaClient()
        self.tmp = tempfile.TemporaryDirectory()
        self.doc = Path(self.tmp.name) / "f.docx"
        _build_doc(self.doc)

    def tearDown(self):
        cli._build_client = self._orig
        self.tmp.cleanup()

    def _run(self, *extra, mode="pro"):
        argv = ["review", str(self.doc), "--mode", mode, "--quiet", *extra]
        return cli.main(argv)

    def test_clean_when_failon_critical_in_max(self):
        # The mock's only CRITICAL is ungrounded; Max mode drops it, so a
        # --fail-on critical gate passes (exit 0). (Recall-first Pro would keep it.)
        self.assertEqual(self._run("--fail-on", "critical", mode="max"), 0)

    def test_fail_on_ungrounded_critical_in_pro(self):
        # Pro keeps the ungrounded CRITICAL (recall-first) → gate trips (exit 1).
        self.assertEqual(self._run("--fail-on", "critical", mode="pro"), 1)

    def test_fail_when_failon_minor(self):
        # A MINOR finding exists, so --fail-on minor must return exit 1.
        self.assertEqual(self._run("--fail-on", "minor"), 1)

    def test_missing_file_returns_2(self):
        cli._build_client = lambda api_key, host: FakeOllamaClient()
        self.assertEqual(cli.main(["review", "does_not_exist.docx", "--quiet"]), 2)


if __name__ == "__main__":
    unittest.main()
