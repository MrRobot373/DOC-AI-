"""Unit tests for the accuracy scorer (no external files / no LLM)."""

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from eval.score import GoldItem, _tokenize, score


class ScorerTests(unittest.TestCase):
    def _gold(self, page, comment):
        return GoldItem(page=page, location="", comment=comment, tokens=_tokenize(comment))

    def test_perfect_match(self):
        gold = [self._gold(5, "output voltage missing units on the regulator")]
        findings = [{"page": "5", "comment": "Regulator output voltage is missing units", "evidence": ""}]
        r = score(findings, gold)
        self.assertEqual(r["matched"], 1)
        self.assertEqual(r["recall"], 1.0)
        self.assertEqual(r["precision"], 1.0)

    def test_no_match_counts_as_miss(self):
        gold = [self._gold(5, "output voltage missing units on the regulator")]
        findings = [{"page": "90", "comment": "completely unrelated typo somewhere", "evidence": ""}]
        r = score(findings, gold)
        self.assertEqual(r["matched"], 0)
        self.assertEqual(r["recall"], 0.0)

    def test_page_mismatch_blocks_match(self):
        # Same words, but pages far apart → not the same issue.
        gold = [self._gold(5, "regulator output voltage missing units")]
        findings = [{"page": "80", "comment": "regulator output voltage missing units", "evidence": ""}]
        r = score(findings, gold)
        self.assertEqual(r["matched"], 0)

    def test_unknown_pages_do_not_block(self):
        gold = [self._gold(None, "regulator output voltage missing units")]
        findings = [{"page": "-", "comment": "regulator output voltage missing units", "evidence": ""}]
        r = score(findings, gold)
        self.assertEqual(r["matched"], 1)


if __name__ == "__main__":
    unittest.main()
