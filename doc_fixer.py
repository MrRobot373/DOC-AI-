"""
Document Fixer Module — Safe Auto-Fix Engine
Applies text-based fixes (spelling, terminology, units) to a COPY of the original document.
NEVER modifies the original. Produces an audit log of all changes.

Safety Rules:
  1. Only applies fixes classified as fix_type="AUTO"
  2. Works at the run level (preserves formatting)
  3. Creates a backup copy before any modifications
  4. Returns a detailed change log
"""

import os
import re
import shutil
from datetime import datetime
from docx import Document


REPORTS_DIR = os.path.join(os.path.dirname(__file__), "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)


def apply_fixes(original_filepath, findings, finding_ids=None):
    """
    Apply auto-fixable findings to a COPY of the original document.
    
    Args:
        original_filepath: Path to the uploaded .docx
        findings: List of finding dicts from review
        finding_ids: Optional list of specific finding IDs to apply.
                     If None, applies ALL auto-fixable findings.
    
    Returns:
        dict with success, fixed_filename, changes_applied, changes_skipped, audit_log
    """
    if not os.path.exists(original_filepath):
        return {"success": False, "error": "Original document not found."}
    
    if not original_filepath.lower().endswith('.docx'):
        return {"success": False, "error": "Auto-fix only supports .docx files."}
    
    # Filter to auto-fixable findings
    auto_findings = []
    for f in findings:
        if f.get("fix_type") != "AUTO":
            continue
        if finding_ids is not None and f.get("id") not in finding_ids:
            continue
        if f.get("status") in ["CLOSED", "IGNORE", "N/A"]:
            continue
        auto_findings.append(f)
    
    if not auto_findings:
        return {
            "success": True,
            "fixed_filename": None,
            "changes_applied": 0,
            "changes_skipped": 0,
            "message": "No auto-fixable findings to apply."
        }
    
    # Create a copy of the document
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = os.path.splitext(os.path.basename(original_filepath))[0]
    # Remove the review_id prefix if present
    clean_name = re.sub(r'^[a-f0-9]{8}_', '', base_name)
    fixed_filename = f"FIXED_{clean_name}_{timestamp}.docx"
    fixed_filepath = os.path.join(REPORTS_DIR, fixed_filename)
    
    shutil.copy2(original_filepath, fixed_filepath)
    
    # Open the copy for editing
    try:
        doc = Document(fixed_filepath)
    except Exception as e:
        return {"success": False, "error": f"Failed to open document copy: {str(e)}"}
    
    # Build replacement map from findings
    replacements = _build_replacement_map(auto_findings)
    
    changes_applied = 0
    changes_skipped = 0
    audit_log = []
    
    # Apply replacements at the run level to preserve formatting
    for para_idx, para in enumerate(doc.paragraphs):
        for replacement in replacements:
            old_text = replacement["old"]
            new_text = replacement["new"]
            finding_id = replacement["finding_id"]
            
            if old_text.lower() in para.text.lower():
                # Try to apply at run level first (preserves formatting)
                applied = _replace_in_runs(para, old_text, new_text)
                if applied:
                    changes_applied += 1
                    audit_log.append({
                        "finding_id": finding_id,
                        "paragraph_index": para_idx,
                        "old_text": old_text,
                        "new_text": new_text,
                        "status": "applied"
                    })
                else:
                    changes_skipped += 1
                    audit_log.append({
                        "finding_id": finding_id,
                        "paragraph_index": para_idx,
                        "old_text": old_text,
                        "new_text": new_text,
                        "status": "skipped_complex"
                    })
    
    # Save the modified copy
    try:
        doc.save(fixed_filepath)
    except Exception as e:
        return {"success": False, "error": f"Failed to save fixed document: {str(e)}"}
    
    # Write audit log
    try:
        audit_path = os.path.join(REPORTS_DIR, f"audit_log_{timestamp}.txt")
        with open(audit_path, "w", encoding="utf-8") as f:
            f.write(f"Auto-Fix Audit Log — {datetime.now().isoformat()}\n")
            f.write(f"Original: {os.path.basename(original_filepath)}\n")
            f.write(f"Fixed: {fixed_filename}\n")
            f.write(f"Changes Applied: {changes_applied}\n")
            f.write(f"Changes Skipped: {changes_skipped}\n")
            f.write("=" * 60 + "\n\n")
            for entry in audit_log:
                f.write(f"Finding #{entry['finding_id']} ({entry['status']}):\n")
                f.write(f"  Old: {entry['old_text']}\n")
                f.write(f"  New: {entry['new_text']}\n")
                f.write(f"  Para Index: {entry['paragraph_index']}\n\n")
    except Exception:
        pass
    
    return {
        "success": True,
        "fixed_filename": fixed_filename,
        "fixed_filepath": fixed_filepath,
        "changes_applied": changes_applied,
        "changes_skipped": changes_skipped,
        "audit_log": audit_log,
    }


def _build_replacement_map(findings):
    """Extract old→new text replacements from findings."""
    replacements = []
    
    for f in findings:
        comment = f.get("comment", "")
        fix = f.get("fix", "")
        finding_id = f.get("id", 0)
        
        # Try to extract old/new from common patterns
        # Pattern 1: "Typo: 'recieve' should be 'receive'"
        match = re.search(r"['\"]([^'\"]+)['\"].*(?:should be|→|->|correct to|replace with)\s*['\"]([^'\"]+)['\"]", comment, re.IGNORECASE)
        if match:
            replacements.append({
                "old": match.group(1),
                "new": match.group(2),
                "finding_id": finding_id,
            })
            continue
        
        # Pattern 2: "Change 'recieve' to 'receive'" (from fix field)
        match = re.search(r"(?:Change|Replace)\s*['\"]([^'\"]+)['\"].*(?:to|with)\s*['\"]([^'\"]+)['\"]", fix, re.IGNORECASE)
        if match:
            replacements.append({
                "old": match.group(1),
                "new": match.group(2),
                "finding_id": finding_id,
            })
            continue
        
        # Pattern 3: "'old' → 'new'" or "'old' -> 'new'"
        match = re.search(r"['\"]([^'\"]+)['\"][\s]*(?:→|->|=>)[\s]*['\"]([^'\"]+)['\"]", comment + " " + fix)
        if match:
            replacements.append({
                "old": match.group(1),
                "new": match.group(2),
                "finding_id": finding_id,
            })
            continue
    
    return replacements


def _replace_in_runs(paragraph, old_text, new_text):
    """
    Replace text within paragraph runs to preserve formatting.
    Returns True if replacement was successful.
    """
    # Simple case: the old text is entirely within one run
    for run in paragraph.runs:
        if old_text in run.text:
            run.text = run.text.replace(old_text, new_text, 1)
            return True
    
    # Case-insensitive single run
    for run in paragraph.runs:
        idx = run.text.lower().find(old_text.lower())
        if idx != -1:
            original = run.text[idx:idx + len(old_text)]
            run.text = run.text[:idx] + new_text + run.text[idx + len(old_text):]
            return True
    
    # Complex case: text spans multiple runs
    # Build a map of character positions to runs
    full_text = paragraph.text
    lower_full = full_text.lower()
    lower_old = old_text.lower()
    
    idx = lower_full.find(lower_old)
    if idx == -1:
        return False
    
    # Find which runs contain the target text
    char_pos = 0
    run_positions = []  # (run_index, start_in_run, end_in_run)
    
    for run_idx, run in enumerate(paragraph.runs):
        run_start = char_pos
        run_end = char_pos + len(run.text)
        
        # Does this run overlap with our target?
        target_start = idx
        target_end = idx + len(old_text)
        
        overlap_start = max(run_start, target_start)
        overlap_end = min(run_end, target_end)
        
        if overlap_start < overlap_end:
            run_positions.append({
                "run_idx": run_idx,
                "run": run,
                "overlap_start": overlap_start - run_start,
                "overlap_end": overlap_end - run_start,
                "is_first": run_start <= target_start,
                "is_last": run_end >= target_end,
            })
        
        char_pos = run_end
    
    if not run_positions:
        return False
    
    # Apply replacement to the first run, clear from subsequent runs
    try:
        first = run_positions[0]
        first_run = first["run"]
        
        # Replace in first run
        before = first_run.text[:first["overlap_start"]]
        after_part = ""
        
        if first == run_positions[-1]:
            # Entire replacement within overlapping runs
            after_part = first_run.text[first["overlap_end"]:]
            first_run.text = before + new_text + after_part
        else:
            first_run.text = before + new_text
            
            # Clear middle runs
            for rp in run_positions[1:-1]:
                rp["run"].text = ""
            
            # Trim last run
            if len(run_positions) > 1:
                last = run_positions[-1]
                last["run"].text = last["run"].text[last["overlap_end"]:]
        
        return True
    except Exception:
        return False
