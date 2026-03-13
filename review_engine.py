"""
Review Engine Module
Takes parsed document data and uses Ollama Cloud API to perform
comprehensive document review across all check categories.
"""

import os
import json
import re
import difflib
from datetime import datetime
from ollama import Client


# All review check categories
REVIEW_CATEGORIES = {
    "GRAMMAR_SPELLING": {
        "name": "Grammar & Spelling",
        "icon": "📝",
        "description": "Grammatical errors, typos, spelling mistakes, sentence construction issues",
    },
    "TERMINOLOGY_CONSISTENCY": {
        "name": "Terminology Consistency",
        "icon": "📖",
        "description": "Inconsistent terms, abbreviations, shortforms used differently across sections",
    },
    "FLOWCHART_DESCRIPTION": {
        "name": "Flowchart ↔ Description",
        "icon": "🔄",
        "description": "Mismatches between flowchart logic and written descriptions",
    },
    "UNITS_CALCULATIONS": {
        "name": "Units & Calculations",
        "icon": "🔬",
        "description": "Incorrect units, missing units, calculation errors, numerical formatting",
    },
    "FORMATTING_ALIGNMENT": {
        "name": "Formatting & Alignment",
        "icon": "📐",
        "description": "Spacing, paragraph alignment, page breaks, margin issues, font inconsistencies",
    },
    "SIGNAL_VARIABLE_NAMING": {
        "name": "Signal & Variable Naming",
        "icon": "🏷️",
        "description": "Pin names, signal names, variable naming inconsistencies",
    },
    "TEST_RESULT_COMPLETENESS": {
        "name": "Test Result Completeness",
        "icon": "✅",
        "description": "Missing criteria, actual measurements, or pass/fail judgments in test results",
    },
    "WAVEFORM_DOCUMENTATION": {
        "name": "Waveform Documentation",
        "icon": "📈",
        "description": "Missing signal names, probe points, legends in oscilloscope waveforms",
    },
    "CROSS_REFERENCE_ACCURACY": {
        "name": "Cross-Reference Accuracy",
        "icon": "🔗",
        "description": "Incorrect figure/table/section references within the document",
    },
    "LOGICAL_CONSISTENCY": {
        "name": "Logical Consistency",
        "icon": "🧠",
        "description": "Logical errors in descriptions, contradictory statements, fault handling mismatches",
    },
    "CONNECTOR_PIN_MAPPING": {
        "name": "Connector & Pin Mapping",
        "icon": "🔌",
        "description": "Missing connector IDs, unclear pin references",
    },
    "MEASUREMENT_RESOLUTION": {
        "name": "Measurement Resolution",
        "icon": "🔍",
        "description": "Debugger/oscilloscope values not readable, insufficient resolution",
    },
}

SEVERITY_LEVELS = {
    "CRITICAL": {"label": "Critical", "color": "#ff4444", "weight": 4},
    "MAJOR": {"label": "Major", "color": "#ff8800", "weight": 3},
    "MINOR": {"label": "Minor", "color": "#ffcc00", "weight": 2},
}


def create_ollama_client(api_key, host="https://ollama.com"):
    """Create an Ollama client with cloud API authentication."""
    client = Client(
        host=host,
        headers={"Authorization": f"Bearer {api_key}"},
    )
    return client


def test_connection(api_key, host="https://ollama.com"):
    """Test connection to Ollama Cloud and return available models."""
    try:
        client = create_ollama_client(api_key, host)
        models = client.list()
        model_names = []
        if hasattr(models, "models"):
            model_names = [m.model for m in models.models]
        elif isinstance(models, dict) and "models" in models:
            model_names = [m.get("name", m.get("model", "unknown")) for m in models["models"]]
        return {"success": True, "models": model_names}
    except Exception as e:
        return {"success": False, "error": str(e)}


def review_document(client, model, parsed_doc, progress_callback=None, review_mode="pro"):
    """
    Perform a comprehensive review of a parsed document.
    """
    findings = []
    
    # Define which categories to use based on mode
    if review_mode == "normal":
        active_categories = ["GRAMMAR_SPELLING", "TERMINOLOGY_CONSISTENCY", "UNITS_CALCULATIONS"]
    else:
        active_categories = list(REVIEW_CATEGORIES.keys())

    try:
        # Step 1: Local checks (no LLM needed)
        if progress_callback:
            progress_callback("Running local formatting checks...")
        
        # Only run local formatting/font checks in Pro mode
        if review_mode == "pro":
            local_findings = _run_local_checks(parsed_doc)
            if local_findings:
                findings.extend(local_findings)

        # Step 2: LLM-powered review of document content in chunks
        from doc_parser import get_document_summary, get_section_chunks

        doc_summary = get_document_summary(parsed_doc)
        chunks = get_section_chunks(parsed_doc, max_chars=5000)
        total_chunks = len(chunks)

        for i, chunk in enumerate(chunks):
            if progress_callback:
                progress_callback(f"Analyzing section {i + 1}/{total_chunks} with AI...")

            try:
                chunk_findings = _review_chunk_with_llm(client, model, chunk, doc_summary, i + 1, active_categories)
                if chunk_findings:
                    findings.extend(chunk_findings)
            except Exception as e:
                pass

        # Step 3: Full-document cross-reference and consistency check
        if progress_callback:
            progress_callback("Checking cross-document consistency...")
        
        # In Normal mode, skip full consistency check
        if review_mode == "pro":
            try:
                consistency_findings = _review_consistency_with_llm(client, model, doc_summary)
                if consistency_findings:
                    findings.extend(consistency_findings)
            except Exception as e:
                findings.append({
                    "category": "CROSS_REFERENCE_ACCURACY",
                    "severity": "MINOR",
                    "page": "-",
                    "section": "ALL",
                    "comment": f"Error during consistency check: {str(e)}",
                    "source": "llm_error"
                })

        # Step 4: Table-specific review
        if parsed_doc.get("tables"):
            if progress_callback:
                progress_callback("Reviewing tables and data...")
            
            if review_mode == "pro":
                try:
                    table_findings = _review_tables_with_llm(client, model, parsed_doc)
                    if table_findings:
                        findings.extend(table_findings)
                except Exception as e:
                    pass
            elif "UNITS_CALCULATIONS" in active_categories:
                # Still check tables for units in normal mode
                try:
                    table_findings = _review_tables_with_llm(client, model, parsed_doc, ["UNITS_CALCULATIONS"])
                    if table_findings:
                        findings.extend(table_findings)
                except Exception as e:
                    pass

    except Exception as e:
        # Catch-all for review logic so we at least return what we found
        findings.append({
            "category": "LOGICAL_CONSISTENCY",
            "severity": "CRITICAL",
            "page": "-",
            "section": "System",
            "comment": f"Critical review engine failure: {str(e)}",
            "source": "system_error"
        })

    # Deduplicate and sort findings safely
    if findings:
        findings = _deduplicate_findings(findings)
        findings.sort(key=lambda f: SEVERITY_LEVELS.get(f.get("severity", "MINOR"), {}).get("weight", 0), reverse=True)

    # Number findings
    for idx, f in enumerate(findings, 1):
        f["id"] = idx

    return findings


def _run_local_checks(parsed_doc):
    """Run checks that don't need LLM - formatting, fonts, spacing, images."""
    findings = []
    fmt = parsed_doc["formatting"]

    # Check for font inconsistencies
    default_font = fmt.get("default_font")
    if default_font:
        font_mismatches = set()
        for section in parsed_doc["sections"]:
            for para in section["paragraphs"]:
                for run in para.get("runs", []):
                    if "font" in run and run["font"] != default_font and run["text"].strip():
                        font_mismatches.add(run["font"])
        if font_mismatches:
            findings.append({
                "category": "FORMATTING_ALIGNMENT",
                "severity": "MINOR",
                "page": "ALL",
                "section": "ALL",
                "comment": (
                    f"Font inconsistency detected. Default font is '{default_font}' but the "
                    f"following fonts are also used: {', '.join(font_mismatches)}. "
                    f"Consider unifying to one font family."
                ),
                "source": "local",
            })

    # Check for font size inconsistencies in body text
    default_size = fmt.get("default_size")
    if default_size:
        size_mismatches = set()
        for section in parsed_doc["sections"]:
            for para in section["paragraphs"]:
                if para["heading_level"]:
                    continue  # Skip headings
                for run in para.get("runs", []):
                    if "size_pt" in run and run["size_pt"] != default_size and run["text"].strip():
                        size_mismatches.add(run["size_pt"])
        if size_mismatches:
            findings.append({
                "category": "FORMATTING_ALIGNMENT",
                "severity": "MINOR",
                "page": "ALL",
                "section": "Body Text",
                "comment": (
                    f"Font size inconsistency in body text. Default size is {default_size}pt "
                    f"but the following sizes are also used: "
                    f"{', '.join(str(s) + 'pt' for s in sorted(size_mismatches))}."
                ),
                "source": "local",
            })

    return findings



def _review_chunk_with_llm(client, model, chunk_text, doc_summary, chunk_num, active_categories):
    """Send a document chunk to the LLM for comprehensive review."""

    cat_list = "\n".join([f"{cid}. {REVIEW_CATEGORIES[cid]['name']}: {REVIEW_CATEGORIES[cid]['description']}" for cid in active_categories if cid in REVIEW_CATEGORIES])

    prompt = f"""You are a senior technical document reviewer for automotive/embedded systems engineering documents.
Review the following section of a technical document and find ALL issues.

IMPORTANT: You must check for ONLY the categories listed below.
CRITICAL: Do NEVER flag errors regarding "incomplete sentences", "abruptly ending text", "garbled text", or "nonsensical values" (e.g., 'xzccccxc', 'PCBDddduiouioDriver'). These are artifacts of the text extraction or test data. Ignore them entirely.

## Review Categories:
{cat_list}

## Document Context:
{doc_summary[:2000]}

## Section to Review (Chunk {chunk_num}):
{chunk_text}

## Output Format:
Return a JSON array of findings. IMPORTANT: Use the (Starts on Page X) or --- [Page X] --- markers from the text to accurately determine and return the "page" number for each finding.
Each finding must be:
{{
  "category": "CATEGORY_ID",
  "severity": "CRITICAL|MAJOR|MINOR",
  "page": "Exact page number (e.g., '14') based on the text markers",
  "section": "section reference",
  "comment": "description of the error/issue",
  "fix": "step-by-step instruction on how to fix this specific error"
}}

Return ONLY the JSON array, no other text. If no issues found, return [].
Example:
[
  {{"category": "GRAMMAR_SPELLING", "severity": "MINOR", "page": "12", "section": "3.1", "comment": "Typo: 'recieve' should be 'receive'", "fix": "Change 'recieve' to 'receive'"}},
  {{"category": "UNITS_CALCULATIONS", "severity": "MAJOR", "page": "15", "section": "4.2", "comment": "Missing unit for voltage value '3.3' - should specify '3.3V' or '3.3 VDC'", "fix": "Add 'V' or 'VDC' after '3.3'"}}
]
"""

    try:
        response = client.chat(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            options={"temperature": 0.1, "num_predict": 4096},
        )

        reply = response["message"]["content"] if isinstance(response, dict) else response.message.content
        return _parse_llm_findings(reply, "llm_chunk")
    except Exception as e:
        return []


def _review_consistency_with_llm(client, model, doc_summary):
    """Check cross-document consistency issues."""

    prompt = f"""You are a senior technical document reviewer. Analyze the overall structure and consistency of this document.

## Full Document Summary:
{doc_summary[:8000]}

## Focus Areas:
1. Cross-reference accuracy (do figure/table/section references point to correct items?)
2. Terminology consistency across the entire document
4. Overall logical flow and completeness
6. Consistent use of abbreviations/shortforms (are they defined on first use?)
7. Consistent formatting patterns across similar sections

CRITICAL: Do NEVER flag errors regarding "incomplete sentences", "abruptly ending text", "garbled text", or "nonsensical values". Ignore extraction artifacts completely.

## Output Format:
Return a JSON array of findings. Each finding must be:
{{
  "category": "CATEGORY_ID",
  "severity": "CRITICAL|MAJOR|MINOR",
  "page": "page number or 'ALL'",
  "section": "section reference or 'ALL'",
  "comment": "description of the inconsistency",
  "fix": "how to resolve the contradiction"
}}

Return ONLY the JSON array. If no issues, return [].
"""

    try:
        response = client.chat(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            options={"temperature": 0.1, "num_predict": 4096},
        )

        reply = response["message"]["content"] if isinstance(response, dict) else response.message.content
        return _parse_llm_findings(reply, "llm_consistency")
    except Exception as e:
        return [{
            "category": "CROSS_REFERENCE_ACCURACY",
            "severity": "MINOR",
            "page": "-",
            "section": "ALL",
            "comment": f"AI consistency check error: {str(e)[:200]}",
            "source": "llm_error",
        }]


def _review_tables_with_llm(client, model, parsed_doc, active_categories=None):
    """Specifically review tables for completeness and formatting."""
    if not parsed_doc["tables"]:
        return []
    
    if not active_categories:
        active_categories = ["UNITS_CALCULATIONS", "TEST_RESULT_COMPLETENESS", "MEASUREMENT_RESOLUTION", "FORMATTING_ALIGNMENT"]

    tables_text = []
    for tbl in parsed_doc["tables"][:15]:  # Limit to first 15 tables
        tbl_name = tbl.get("name", f"Table {tbl['index'] + 1}")
        rows_str = "\n".join(
            f"  Row {i}: {' | '.join(c[:150] for c in row)}"
            for i, row in enumerate(tbl["rows"][:50])
        )
        tables_text.append(f"--- {tbl_name} ({tbl['num_rows']}×{tbl['num_cols']}) ---\n{rows_str}")

    prompt = f"""You are reviewing TABLES in a technical document. Check each table for:
1. Missing headers or unclear column names
2. Empty cells that should have values
3. Inconsistent units across rows
4. Test results without pass/fail criteria, actual measurements, or judgments
5. Numerical values without units
6. Calculation errors or suspicious values
7. Consistent formatting across rows

IMPORTANT: Do NOT report errors stating that a table is "incomplete" or "truncated". Assume data may legitimately continue on another page.
CRITICAL: Do NEVER flag errors regarding "garbled text" or "nonsensical values" (e.g., 'xzccccxc', 'PCBDddduiouioDriver'). Ignore placeholder or extraction artifact text completely.

## Tables:
{chr(10).join(tables_text)}

## Output Format:
Return a JSON array of findings. Each finding must be:
{{
  "category": "CATEGORY_ID",
  "severity": "CRITICAL|MAJOR|MINOR",
  "page": "-",
  "section": "Exact Table Name (from the --- Name --- marker)",
  "comment": "detailed description",
  "fix": "step-by-step instruction on how to fix this specific error"
}}

Return ONLY the JSON array. If no issues, return [].
"""

    try:
        response = client.chat(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            options={"temperature": 0.1, "num_predict": 4096},
        )
        reply = response["message"]["content"] if isinstance(response, dict) else response.message.content
        return _parse_llm_findings(reply, "llm_tables")
    except Exception as e:
        return []


def _parse_llm_findings(llm_response, source="llm"):
    """Parse LLM response into structured findings."""
    # Try to extract JSON from response
    text = llm_response.strip()

    # Remove markdown code blocks if present
    if "```json" in text:
        text = text.split("```json")[1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[1].split("```")[0].strip()

    # Find the JSON array
    start = text.find("[")
    end = text.rfind("]")
    if start == -1 or end == -1:
        return []

    json_text = text[start : end + 1]

    try:
        findings = json.loads(json_text)
        if not isinstance(findings, list):
            return []

        # Validate and normalize each finding
        valid_findings = []
        for f in findings:
            if not isinstance(f, dict):
                continue
            category = f.get("category", "GRAMMAR_SPELLING")
            if category not in REVIEW_CATEGORIES:
                category = "GRAMMAR_SPELLING"

            severity = f.get("severity", "MINOR").upper()
            if severity == "SUGGESTION":
                continue  # Skip suggestions entirely as requested
            if severity not in SEVERITY_LEVELS:
                severity = "MINOR"

            valid_findings.append({
                "category": category,
                "severity": severity,
                "page": str(f.get("page", "-")),
                "section": str(f.get("section", "-")),
                "comment": str(f.get("comment", "")),
                "fix": str(f.get("fix", "Review the content and apply standard technical writing guidelines.")),
                "source": source,
            })

        return valid_findings
    except json.JSONDecodeError:
        return []


def _deduplicate_findings(findings):
    """Remove duplicate or very similar findings."""
    unique = []
    for f in findings:
        is_duplicate = False
        for u in unique:
            # If category is the same, and they are either in the same section or same page, check similarity
            if f.get('category') == u.get('category'):
                # Check comment similarity (70% match is considered a duplicate)
                ratio = difflib.SequenceMatcher(None, f.get('comment', '').lower(), u.get('comment', '').lower()).ratio()
                if ratio > 0.7:
                    is_duplicate = True
                    break
        if not is_duplicate:
            unique.append(f)
    return unique
