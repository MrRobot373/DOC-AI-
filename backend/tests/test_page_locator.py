"""Tests for the pure page-anchoring logic (no real PDF renderer needed)."""

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from page_locator import assign_pages


class PageAnchorTests(unittest.TestCase):
    def _doc(self):
        return {
            "headings": [
                {"text": "Introduction", "index": 0, "page": 99},
                {"text": "High level requirements", "index": 2, "page": 99},
            ],
            "sections": [
                {"heading": "Introduction", "page": 99, "paragraphs": [
                    {"index": 0, "text": "Introduction", "page": 99},
                    {"index": 1, "text": "Scope of this document is the power supply.", "page": 99},
                ]},
                {"heading": "High level requirements", "page": 99, "paragraphs": [
                    {"index": 2, "text": "High level requirements", "page": 99},
                    {"index": 3, "text": "The output voltage is 3.3 V nominal.", "page": 99},
                ]},
            ],
            "statistics": {},
        }

    def test_pages_anchored_from_pdf_text(self):
        # page 1 holds the intro; the requirements section starts on page 4.
        page_texts = [
            "introduction scope of this document is the power supply.",
            "filler page two", "filler page three",
            "high level requirements the output voltage is 3.3 v nominal.",
        ]
        parsed = assign_pages(self._doc(), page_texts)
        pages = {h["text"]: h["page"] for h in parsed["headings"]}
        self.assertEqual(pages["Introduction"], 1)
        self.assertEqual(pages["High level requirements"], 4)
        self.assertEqual(parsed["statistics"]["total_pages"], 4)

    def test_pages_are_monotonic(self):
        page_texts = [
            "introduction scope of this document is the power supply.",
            "high level requirements the output voltage is 3.3 v nominal.",
        ]
        parsed = assign_pages(self._doc(), page_texts)
        seen = [p["page"] for s in parsed["sections"] for p in s["paragraphs"]]
        self.assertEqual(seen, sorted(seen), "Paragraph pages must never decrease")

    def test_empty_page_texts_is_noop(self):
        doc = self._doc()
        before = [p["page"] for s in doc["sections"] for p in s["paragraphs"]]
        assign_pages(doc, [])
        after = [p["page"] for s in doc["sections"] for p in s["paragraphs"]]
        self.assertEqual(before, after)


if __name__ == "__main__":
    unittest.main()
