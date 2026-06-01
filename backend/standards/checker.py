"""
Standards compliance checker.

Deterministic, high-precision structural checks driven by JSON rule-packs:
  - required-section presence (is the Hazard Analysis / SIL / etc. present at all?)
  - FMEA table column completeness (does the failure-mode table have cause/effect/
    detection/mitigation columns, and are they populated?)
  - naming conventions (regex, e.g. AUTOSAR component/port names)

Returns findings in the standard engine shape (category/severity/comment/...), plus
a compliance checklist structure the report can render as a "Compliance" sheet.
"""

from __future__ import annotations

import json
import os
import re

RULES_DIR = os.path.dirname(os.path.abspath(__file__))

# Categories surfaced by the standards checker (registered in review_engine too).
STANDARDS_CATEGORIES = {
    "ISO26262_COMPLIANCE": {"name": "ISO 26262 Compliance", "icon": "🛡️"},
    "IEC61508_COMPLIANCE": {"name": "IEC 61508 Compliance", "icon": "⚙️"},
    "AUTOSAR_NAMING": {"name": "AUTOSAR Naming", "icon": "🔤"},
    "FMEA_COVERAGE": {"name": "FMEA Coverage", "icon": "🧯"},
    "TRACEABILITY": {"name": "Traceability", "icon": "🔗"},
}


def available_standards():
    """List the rule-packs that ship with the app (id + display name)."""
    out = []
    for fn in sorted(os.listdir(RULES_DIR)):
        if fn.endswith(".json"):
            try:
                with open(os.path.join(RULES_DIR, fn), encoding="utf-8") as fh:
                    pack = json.load(fh)
                out.append({"id": pack["id"], "name": pack.get("name", pack["id"])})
            except Exception:
                pass
    return out


def _load_pack(standard_id):
    path = os.path.join(RULES_DIR, f"{standard_id}.json")
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def _check_required_sections(pack, raw_lower, checklist):
    findings = []
    cat = pack["category"]
    for req in pack.get("required_sections", []):
        present = any(p.lower() in raw_lower for p in req.get("patterns", []))
        checklist.append({"standard": pack["name"], "element": req["label"],
                          "status": "Present" if present else "Missing"})
        if not present:
            findings.append({
                "category": cat, "severity": "MAJOR", "page": "-",
                "section": "Standards compliance",
                "comment": f"{pack['name']}: required element '{req['label']}' was not found in the document.",
                "fix": f"Add a '{req['label']}' section/content to satisfy {pack['name']}.",
                "evidence": req["label"],
                "source": "standards", "fix_type": "MANUAL",
            })
    return findings


def _check_fmea_tables(pack, parsed_doc, checklist):
    cfg = pack.get("fmea_tables") or {}
    detect = [h.lower() for h in cfg.get("detect_headers", [])]
    required = [c.lower() for c in cfg.get("required_columns", [])]
    if not detect or not required:
        return []
    findings = []
    found_any = False
    for tbl in parsed_doc.get("tables", []):
        rows = tbl.get("rows", [])
        if not rows:
            continue
        header = " | ".join(rows[0]).lower()
        if not any(d in header for d in detect):
            continue
        found_any = True
        missing = [c for c in required if c not in header]
        name = tbl.get("name", f"Table {tbl.get('index', 0) + 1}")
        if missing:
            findings.append({
                "category": pack["category"], "severity": "MAJOR", "page": "-",
                "section": name,
                "comment": f"{pack['name']}: FMEA table '{name}' is missing required column(s): {', '.join(missing)}.",
                "fix": f"Add the missing FMEA column(s): {', '.join(missing)}.",
                "evidence": rows[0][0] if rows[0] else name,
                "source": "standards", "fix_type": "MANUAL",
            })
    if detect and not found_any:
        checklist.append({"standard": pack["name"], "element": "FMEA table with required columns",
                          "status": "Missing"})
    return findings


def _check_naming(pack, parsed_doc, checklist):
    findings = []
    raw = parsed_doc.get("raw_text", "")
    for rule in pack.get("naming", []):
        try:
            find_re = re.compile(rule["find"])
            must_re = re.compile(rule["must_match"])
        except re.error:
            continue
        violations = []
        for m in find_re.finditer(raw):
            name = m.group(1)
            if name and not must_re.match(name):
                violations.append(name)
        for v in sorted(set(violations))[:10]:
            findings.append({
                "category": pack["category"], "severity": "MINOR", "page": "-",
                "section": "Naming conventions",
                "comment": f"{pack['name']}: '{v}' does not follow the convention — {rule['label']}.",
                "fix": rule.get("hint", "Rename to match the convention."),
                "evidence": v,
                "source": "standards", "fix_type": "MANUAL",
            })
    return findings


def run_standards_checks(parsed_doc, standards_selected):
    """
    Run the selected standards rule-packs against a parsed document.

    Returns (findings, checklist) where checklist is a list of
    {standard, element, status} rows for the report's Compliance sheet.
    """
    findings = []
    checklist = []
    raw_lower = (parsed_doc.get("raw_text", "") or "").lower()
    for sid in standards_selected or []:
        pack = _load_pack(sid)
        if not pack:
            continue
        findings.extend(_check_required_sections(pack, raw_lower, checklist))
        findings.extend(_check_fmea_tables(pack, parsed_doc, checklist))
        findings.extend(_check_naming(pack, parsed_doc, checklist))
    return findings, checklist
