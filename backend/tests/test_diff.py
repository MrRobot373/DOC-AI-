"""Tests for version-comparison diff_findings()."""

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from review_engine import diff_findings


def _f(cat, comment, evidence=""):
    return {"category": cat, "comment": comment, "evidence": evidence, "severity": "MAJOR"}


class DiffTests(unittest.TestCase):
    def test_new_fixed_unchanged(self):
        old = [
            _f("GRAMMAR_SPELLING", "Typo 'recieve' should be 'receive'", "recieve"),
            _f("UNITS_CALCULATIONS", "Missing unit on output voltage", "output voltage"),
        ]
        new = [
            _f("UNITS_CALCULATIONS", "Missing unit on output voltage", "output voltage"),  # unchanged
            _f("TOC_VALIDATION", "Heading 3.2 missing from TOC", "3.2"),                    # new
        ]
        d = diff_findings(old, new)
        self.assertEqual(len(d["unchanged"]), 1)
        self.assertEqual(len(d["new"]), 1)
        self.assertEqual(len(d["fixed"]), 1)
        self.assertEqual(d["new"][0]["category"], "TOC_VALIDATION")
        self.assertEqual(d["fixed"][0]["category"], "GRAMMAR_SPELLING")

    def test_all_fixed_when_new_empty(self):
        old = [_f("GRAMMAR_SPELLING", "x", "x")]
        d = diff_findings(old, [])
        self.assertEqual(len(d["fixed"]), 1)
        self.assertEqual(len(d["new"]), 0)

    def test_all_new_when_old_empty(self):
        new = [_f("GRAMMAR_SPELLING", "x", "x")]
        d = diff_findings([], new)
        self.assertEqual(len(d["new"]), 1)
        self.assertEqual(len(d["fixed"]), 0)


if __name__ == "__main__":
    unittest.main()
