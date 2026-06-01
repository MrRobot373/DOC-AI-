"""Standards compliance checker tests (deterministic, no LLM)."""

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from standards.checker import run_standards_checks, available_standards


class StandardsTests(unittest.TestCase):
    def test_available_standards_lists_packs(self):
        ids = {s["id"] for s in available_standards()}
        self.assertTrue({"iso26262", "iec61508", "autosar", "fmea"}.issubset(ids))

    def test_missing_iso26262_elements_flagged(self):
        parsed = {"raw_text": "This document describes a power supply. No safety content here.",
                  "tables": [], "sections": []}
        findings, checklist = run_standards_checks(parsed, ["iso26262"])
        # All required ISO 26262 elements are absent → each flagged
        self.assertTrue(any("Safety Goals" in f["comment"] for f in findings))
        self.assertTrue(all(f["category"] == "ISO26262_COMPLIANCE" for f in findings))
        self.assertTrue(any(c["status"] == "Missing" for c in checklist))

    def test_present_elements_not_flagged(self):
        parsed = {
            "raw_text": ("Safety Goals are defined. ASIL B classification applies. "
                         "Hazard Analysis and Risk Assessment (HARA) was performed. "
                         "Safety mechanisms are listed. Safety requirements are traced. "
                         "An FMEA failure mode analysis is included."),
            "tables": [], "sections": [],
        }
        findings, checklist = run_standards_checks(parsed, ["iso26262"])
        # No missing required-SECTION findings expected.
        missing = [f for f in findings if "required element" in f["comment"]]
        self.assertEqual(missing, [])
        # All required-section checklist rows are Present (the optional FMEA-table
        # row may still be Missing since there are no tables in this fixture).
        section_rows = [c for c in checklist if c["element"] != "FMEA table with required columns"]
        self.assertTrue(all(c["status"] == "Present" for c in section_rows))

    def test_fmea_table_missing_columns_flagged(self):
        parsed = {
            "raw_text": "FMEA included.",
            "sections": [],
            "tables": [{
                "index": 0, "name": "FMEA Table", "num_rows": 2, "num_cols": 2,
                "rows": [["Failure Mode", "Severity"], ["open circuit", "high"]],
            }],
        }
        findings, _ = run_standards_checks(parsed, ["fmea"])
        # 'cause', 'effect', 'detection', 'occurrence' columns are missing
        self.assertTrue(any("missing required column" in f["comment"].lower() for f in findings))


if __name__ == "__main__":
    unittest.main()
