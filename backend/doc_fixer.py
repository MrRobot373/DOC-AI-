"""
Document Fixer — Tracked Changes + Word Comments (production-safe version)

Instead of silently rewriting text, this module produces a .docx where:
  • AUTO findings  → inserted as Word TRACKED CHANGES (the engineer sees
                     red/green strikethrough+insert marks; clicks Accept/Reject)
  • ALL findings   → inserted as Word COMMENTS anchored to the evidence quote
                     (visible as comment balloons when opened in Word/LibreOffice)

The original document is NEVER modified. A copy is created and returned.

Why this is better than silent edits:
  - Nothing changes behind the reviewer's back.
  - The engineer can Accept / Reject each change in Word in one click.
  - Comment balloons show the exact AI reasoning at the problem location.
  - A malformed tracked-change is visible and harmless; a silent wrong edit
    could corrupt a client deliverable.
"""

from __future__ import annotations

import copy
import os
import re
import shutil
from datetime import datetime
from typing import Optional

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

REPORTS_DIR = os.path.join(os.path.dirname(__file__), "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)

AUTHOR = "DOC-AI"
REVIEW_DATE = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")


# ── Helpers ──────────────────────────────────────────────────────────────────

def _make_elem(tag: str, **attrs) -> etree._Element:
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn(k) if ":" in k else k, v)
    return el


def _next_id(doc: Document) -> int:
    """Generate a unique revision/comment id within the document."""
    used: set[int] = set()
    for el in doc.element.body.iter():
        for attr in ("w:id",):
            v = el.get(qn(attr))
            if v and v.isdigit():
                used.add(int(v))
    return max(used, default=0) + 1


# ── Tracked Change helpers ───────────────────────────────────────────────────

def _make_del_run(run_el: etree._Element, rev_id: int) -> etree._Element:
    """Wrap an existing run element in a w:del (tracked deletion)."""
    w_del = _make_elem("w:del", **{"w:id": str(rev_id), "w:author": AUTHOR, "w:date": REVIEW_DATE})
    r_copy = copy.deepcopy(run_el)
    # Change w:t → w:delText
    for t in r_copy.findall(qn("w:t")):
        t.tag = qn("w:delText")
    w_del.append(r_copy)
    return w_del


def _make_ins_run(new_text: str, rpr_el: Optional[etree._Element], rev_id: int) -> etree._Element:
    """Create a w:ins (tracked insertion) run with the given text."""
    w_ins = _make_elem("w:ins", **{"w:id": str(rev_id), "w:author": AUTHOR, "w:date": REVIEW_DATE})
    r = _make_elem("w:r")
    if rpr_el is not None:
        r.append(copy.deepcopy(rpr_el))
    t = _make_elem("w:t")
    t.text = new_text
    if new_text and (new_text[0] == " " or new_text[-1] == " "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    w_ins.append(r)
    return w_ins


def _apply_tracked_change(para_el: etree._Element, old_text: str, new_text: str, rev_id: int) -> bool:
    """
    Find `old_text` inside a paragraph element, replace it with a
    tracked deletion + tracked insertion. Returns True on success.

    Works at the run level; if text spans multiple runs it falls back to
    rebuilding the paragraph with a single replacement run (rare case).
    """
    runs = para_el.findall(qn("w:r"))
    # Build full paragraph text and a character→run map
    char_map: list[tuple[etree._Element, int]] = []  # (run_el, char_index_in_run)
    for r in runs:
        t = r.find(qn("w:t"))
        if t is None or not t.text:
            continue
        for i, _ in enumerate(t.text):
            char_map.append((r, i))

    full = "".join(
        (r.find(qn("w:t")).text or "") for r in runs if r.find(qn("w:t")) is not None
    )
    idx = full.lower().find(old_text.lower())
    if idx == -1:
        return False

    start, end = idx, idx + len(old_text)

    # Find the unique runs spanning [start, end)
    involved = []
    for ci in range(start, end):
        r_el, _ = char_map[ci]
        if not involved or involved[-1] is not r_el:
            involved.append(r_el)

    if not involved:
        return False

    first_run = involved[0]
    rpr = first_run.find(qn("w:rPr"))

    # For each involved run: split out the target text and replace with del+ins
    # Simple case: entire replacement is within a single run
    if len(involved) == 1:
        t_el = first_run.find(qn("w:t"))
        run_text = t_el.text or ""
        run_idx = full[:start].count("") - sum(
            len(r.find(qn("w:t")).text or "") for r in runs
            if r.find(qn("w:t")) is not None and list(runs).index(r) < list(runs).index(first_run)
        )
        # Simpler: use the char_map
        start_in_run = char_map[start][1]
        end_in_run = char_map[end - 1][1] + 1

        before = run_text[:start_in_run]
        after = run_text[end_in_run:]

        # Build: before_run | w:del | w:ins | after_run
        parent = first_run.getparent()
        pos = list(parent).index(first_run)

        if before:
            br = copy.deepcopy(first_run)
            br.find(qn("w:t")).text = before
            parent.insert(pos, br)
            pos += 1

        parent.insert(pos, _make_del_run(first_run, rev_id))
        pos += 1
        parent.insert(pos, _make_ins_run(new_text, rpr, rev_id + 1))
        pos += 1

        if after:
            ar = copy.deepcopy(first_run)
            ar.find(qn("w:t")).text = after
            parent.insert(pos, ar)

        parent.remove(first_run)
        return True

    # Multi-run case: just delete all involved runs and insert one new tracked insertion
    parent = involved[0].getparent()
    pos = list(parent).index(involved[0])
    for r in involved:
        parent.insert(pos, _make_del_run(r, rev_id))
        rev_id += 1
        pos += 1
        parent.remove(r)
    parent.insert(pos, _make_ins_run(new_text, rpr, rev_id))
    return True


# ── Comment helpers ───────────────────────────────────────────────────────────

def _ensure_comments_part(doc: Document) -> etree._Element:
    """Return (or create) the word/comments.xml root element."""
    try:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        comments_part = doc.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        )
        return comments_part._element
    except Exception:
        pass

    # Create from scratch
    nsmap = {
        "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    }
    root = etree.Element(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comments",
        nsmap={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
    )
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI
    CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
    RTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    part = Part(PackURI("/word/comments.xml"), CT, etree.tostring(root, xml_declaration=True, encoding="UTF-8"), doc.part.package)
    part._element = root
    doc.part.relate_to(part, RTYPE)
    return root


def _add_comment(doc: Document, para_el: etree._Element, anchor_text: str,
                 comment_text: str, comment_id: int) -> bool:
    """
    Add a Word comment anchored to the first occurrence of `anchor_text` in
    the paragraph. Falls back to attaching at paragraph level if not found.
    """
    comments_root = _ensure_comments_part(doc)

    # Build the <w:comment> element
    comment_el = _make_elem("w:comment", **{
        "w:id": str(comment_id),
        "w:author": AUTHOR,
        "w:date": REVIEW_DATE,
        "w:initials": "AI",
    })
    cp = _make_elem("w:p")
    cr = _make_elem("w:r")
    ct = _make_elem("w:t")
    ct.text = comment_text
    if comment_text and (comment_text[0] == " " or comment_text[-1] == " "):
        ct.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    cr.append(ct)
    cp.append(cr)
    comment_el.append(cp)
    comments_root.append(comment_el)

    # Find anchor position in paragraph
    runs = para_el.findall(qn("w:r"))
    full = "".join((r.find(qn("w:t")).text or "") for r in runs if r.find(qn("w:t")) is not None)
    idx = full.lower().find((anchor_text or "")[:50].lower()) if anchor_text else -1

    start_el = _make_elem("w:commentRangeStart", **{"w:id": str(comment_id)})
    end_el = _make_elem("w:commentRangeEnd", **{"w:id": str(comment_id)})
    ref_r = _make_elem("w:r")
    ref = _make_elem("w:commentReference", **{"w:id": str(comment_id)})
    ref_r.append(ref)

    if idx != -1:
        # Find the run containing the start of the anchor
        pos = 0
        for r in runs:
            t = r.find(qn("w:t"))
            rlen = len(t.text or "") if t is not None else 0
            if pos + rlen > idx:
                parent = r.getparent()
                r_pos = list(parent).index(r)
                parent.insert(r_pos, start_el)
                parent.insert(r_pos + 2, end_el)
                parent.insert(r_pos + 3, ref_r)
                return True
            pos += rlen

    # Fallback: append at end of paragraph
    para_el.append(start_el)
    para_el.append(end_el)
    para_el.append(ref_r)
    return True


# ── Public API ────────────────────────────────────────────────────────────────

def apply_fixes(original_filepath: str, findings: list, finding_ids=None) -> dict:
    """
    Apply AUTO findings as tracked changes and insert ALL findings as Word
    comments anchored to the evidence text in a copy of the document.

    Args:
        original_filepath: Path to the uploaded .docx
        findings:          List of finding dicts from the review engine
        finding_ids:       Optional list of specific finding IDs to apply.
                           None = apply all AUTO findings + comment all findings.

    Returns:
        {success, fixed_filename, changes_applied, changes_skipped, comments_added}
    """
    if not os.path.exists(original_filepath):
        return {"success": False, "error": "Original document not found."}
    if not original_filepath.lower().endswith(".docx"):
        return {"success": False, "error": "Auto-fix only supports .docx files."}

    # Create a copy — never touch the original
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    clean_name = re.sub(r"^[a-f0-9]{8}_", "", os.path.splitext(os.path.basename(original_filepath))[0])
    fixed_filename = f"FIXED_{clean_name}_{timestamp}.docx"
    fixed_path = os.path.join(REPORTS_DIR, fixed_filename)
    shutil.copy2(original_filepath, fixed_path)

    try:
        doc = Document(fixed_path)
    except Exception as e:
        return {"success": False, "error": f"Failed to open document copy: {e}"}

    # Determine which findings to process
    active = []
    for f in findings:
        if finding_ids is not None and f.get("id") not in finding_ids:
            continue
        if f.get("status") in ("CLOSED", "IGNORE", "N/A"):
            continue
        active.append(f)

    auto_findings = [f for f in active if f.get("fix_type") == "AUTO"]

    # Build replacement map for AUTO findings
    replacements = _build_replacement_map(auto_findings)

    changes_applied = 0
    changes_skipped = 0
    comments_added = 0
    rev_id = _next_id(doc)

    # ── Pass 1: Tracked changes for AUTO findings ──────────────────────────
    paragraphs = list(doc.paragraphs)
    for para in paragraphs:
        for repl in replacements:
            if repl["old"].lower() not in para.text.lower():
                continue
            ok = _apply_tracked_change(para._element, repl["old"], repl["new"], rev_id)
            rev_id += 2
            if ok:
                changes_applied += 1
            else:
                changes_skipped += 1

    # ── Pass 2: Word comments for ALL active findings ─────────────────────
    # Build a normalized text index: para_index → paragraph object
    para_index = {p._element: i for i, p in enumerate(paragraphs)}

    for finding in active:
        evidence = (finding.get("evidence") or finding.get("comment") or "")[:80]
        if not evidence:
            continue
        comment_text = _build_comment_text(finding)
        placed = False

        # Try to anchor to the paragraph containing the evidence
        for para in paragraphs:
            if evidence.lower()[:30] in para.text.lower():
                _add_comment(doc, para._element, evidence, comment_text, rev_id)
                rev_id += 1
                comments_added += 1
                placed = True
                break

        # Fallback: anchor to first paragraph
        if not placed and paragraphs:
            _add_comment(doc, paragraphs[0]._element, "", comment_text, rev_id)
            rev_id += 1
            comments_added += 1

    try:
        doc.save(fixed_path)
    except Exception as e:
        return {"success": False, "error": f"Failed to save fixed document: {e}"}

    return {
        "success": True,
        "fixed_filename": fixed_filename,
        "changes_applied": changes_applied,
        "changes_skipped": changes_skipped,
        "comments_added": comments_added,
        "message": (
            f"{changes_applied} tracked change(s) applied, "
            f"{comments_added} comment(s) inserted. "
            f"Open the document in Word to Accept/Reject changes."
        ),
    }


def _build_comment_text(finding: dict) -> str:
    sev = finding.get("severity", "")
    cat = finding.get("category", "").replace("_", " ").title()
    conf = finding.get("confidence")
    conf_str = f" (confidence {conf:.0%})" if isinstance(conf, (int, float)) else ""
    comment = finding.get("comment", "")
    fix = finding.get("fix", "")
    lines = [f"[{sev} — {cat}{conf_str}]", comment]
    if fix:
        lines.append(f"Fix: {fix}")
    return "\n".join(lines)


def _build_replacement_map(auto_findings: list) -> list:
    """Extract old→new text pairs from AUTO findings using comment/fix heuristics."""
    replacements = []
    for f in auto_findings:
        comment = f.get("comment", "")
        fix = f.get("fix", "")
        finding_id = f.get("id", 0)

        # Parse old->new from patterns like: 'x' should be 'y' / 'x' -> 'y'
        for text in (comment, fix):
            m = __import__('re').search(
                r'["\x27]((?:[^"\x27]+))["\x27]'
                r'\s*(?:should be|->|=>|correct to|replace with|with|changed to|to)\s*'
                r'["\x27]((?:[^"\x27]+))["\x27]',
                text, __import__('re').IGNORECASE | __import__('re').DOTALL,
            )
            if m and m.group(1).strip() and m.group(2).strip():
                replacements.append({
                    'old': m.group(1).strip(),
                    'new': m.group(2).strip(),
                    'finding_id': finding_id,
                })
                break
    return replacements
