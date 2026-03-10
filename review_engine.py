"""
Review Engine Module
Takes parsed document data and uses Ollama Cloud API to perform
comprehensive document review across all check categories.
"""

import os
import json
import re
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
    "MISSING_CONTENT": {
        "name": "Missing Content",
        "icon": "⚠️",
        "description": "Sections with only figures/tables without explanatory text, missing descriptions",
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
    "SUGGESTION": {"label": "Suggestion", "color": "#44aaff", "weight": 1},
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


def review_document(client, model, parsed_doc, progress_callback=None):
    """
    Perform comprehensive document review using Ollama LLM.
    Returns a list of findings.
    """
    findings = []

    try:
        # Step 1: Local checks (no LLM needed)
        if progress_callback:
            progress_callback("Running local formatting checks...")
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
                chunk_findings = _review_chunk_with_llm(client, model, chunk, doc_summary, i + 1)
                if chunk_findings:
                    findings.extend(chunk_findings)
            except Exception as e:
                findings.append({
                    "category": "MISSING_CONTENT",
                    "severity": "MINOR",
                    "page": "-",
                    "section": f"Chunk {i+1}",
                    "comment": f"Error parsing this section: {str(e)}",
                    "source": "llm_error"
                })

        # Step 3: Full-document cross-reference and consistency check
        if progress_callback:
            progress_callback("Running cross-document consistency check...")
        
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
        if parsed_doc.get("tables") and progress_callback:
            progress_callback("Reviewing tables...")
            try:
                table_findings = _review_tables_with_llm(client, model, parsed_doc)
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

    # Check for empty sections
    for section in parsed_doc["sections"]:
        texts = [p["text"] for p in section["paragraphs"] if p["text"]]
        has_image = any(p["has_image"] for p in section["paragraphs"])
        if not texts and has_image:
            findings.append({
                "category": "MISSING_CONTENT",
                "severity": "MAJOR",
                "page": "-",
                "section": section["heading"],
                "comment": (
                    f"Section '{section['heading']}' contains only images/figures without any "
                    f"explanatory text. Add descriptions explaining the content."
                ),
                "source": "local",
            })
        elif len(texts) <= 1 and has_image:
            findings.append({
                "category": "MISSING_CONTENT",
                "severity": "MINOR",
                "page": "-",
                "section": section["heading"],
                "comment": (
                    f"Section '{section['heading']}' has minimal text alongside images. "
                    f"Consider adding more detailed descriptions."
                ),
                "source": "local",
            })

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


def _review_chunk_with_llm(client, model, chunk_text, doc_summary, chunk_num):
    """Send a document chunk to the LLM for comprehensive review."""

    prompt = f"""You are a senior technical document reviewer for automotive/embedded systems engineering documents.
Review the following section of a technical document and find ALL issues.

IMPORTANT: You must check for EVERY category listed below. Be thorough and flag everything, even minor issues.

## Review Categories:
1. GRAMMAR_SPELLING: Grammar errors, typos, spelling mistakes, sentence construction problems
3. TERMINOLOGY_CONSISTENCY: Same concept referred to with different terms (e.g., "Flying cap" vs "Flying Capacitor")
4. FLOWCHART_DESCRIPTION: Descriptions that contradict or don't match referenced flowcharts/diagrams
6. UNITS_CALCULATIONS: Missing units, wrong units, calculation errors, incorrect numerical values
7. FORMATTING_ALIGNMENT: Spacing issues, inconsistent formatting, alignment problems
8. SIGNAL_VARIABLE_NAMING: Inconsistent signal/variable names (e.g., "Vdclink" vs "Vdc")
9. TEST_RESULT_COMPLETENESS: Test results missing expected values, actual measurements, or pass/fail
10. WAVEFORM_DOCUMENTATION: Waveform references missing signal names, probe points, legends
11. CROSS_REFERENCE_ACCURACY: Wrong figure/section/table references
12. LOGICAL_CONSISTENCY: Logical errors, contradictions, fault handling mismatches
13. MISSING_CONTENT: Sections with insufficient description
14. CONNECTOR_PIN_MAPPING: Connector references without proper IDs
15. MEASUREMENT_RESOLUTION: Measurement values that seem unclear or at wrong precision

## Document Context:
{doc_summary[:2000]}

## Section to Review (Chunk {chunk_num}):
{chunk_text}

## Output Format:
Return a JSON array of findings. IMPORTANT: Use the (Starts on Page X) or --- [Page X] --- markers from the text to accurately determine and return the "page" number for each finding.
Each finding must be:
{{
  "category": "CATEGORY_ID",
  "severity": "CRITICAL|MAJOR|MINOR|SUGGESTION",
  "page": "Exact page number (e.g., '14') based on the text markers",
  "section": "section reference",
  "comment": "detailed description of the issue and suggested fix"
}}

Return ONLY the JSON array, no other text. If no issues found, return [].
Example:
[
  {{"category": "GRAMMAR_SPELLING", "severity": "MINOR", "page": "12", "section": "3.1", "comment": "Typo: 'recieve' should be 'receive'"}},
  {{"category": "UNITS_CALCULATIONS", "severity": "MAJOR", "page": "15", "section": "4.2", "comment": "Missing unit for voltage value '3.3' - should specify '3.3V' or '3.3 VDC'"}}
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
        return [{
            "category": "MISSING_CONTENT",
            "severity": "SUGGESTION",
            "page": "-",
            "section": f"Chunk {chunk_num}",
            "comment": f"AI review error for this section: {str(e)[:200]}",
            "source": "llm_error",
        }]


def _review_consistency_with_llm(client, model, doc_summary):
    """Check cross-document consistency issues."""

    prompt = f"""You are a senior technical document reviewer. Analyze the overall structure and consistency of this document.

## Full Document Summary:
{doc_summary[:8000]}

## Focus Areas:
1. Cross-reference accuracy (do figure/table/section references point to correct items?)
2. Terminology consistency across the entire document
4. Overall logical flow and completeness
5. Missing sections or descriptions
6. Consistent use of abbreviations/shortforms (are they defined on first use?)
7. Consistent formatting patterns across similar sections

## Output Format:
Return a JSON array of findings. Each finding must be:
{{
  "category": "CATEGORY_ID",
  "severity": "CRITICAL|MAJOR|MINOR|SUGGESTION",
  "page": "page number or 'ALL'",
  "section": "section reference or 'ALL'",
  "comment": "detailed description of the issue"
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
            "severity": "SUGGESTION",
            "page": "-",
            "section": "ALL",
            "comment": f"AI consistency check error: {str(e)[:200]}",
            "source": "llm_error",
        }]


def _review_tables_with_llm(client, model, parsed_doc):
    """Specifically review tables for completeness and formatting."""
    if not parsed_doc["tables"]:
        return []

    tables_text = []
    for tbl in parsed_doc["tables"][:10]:  # Limit to first 10 tables
        rows_str = "\n".join(
            f"  Row {i}: {' | '.join(c[:60] for c in row)}"
            for i, row in enumerate(tbl["rows"][:10])
        )
        tables_text.append(f"Table {tbl['index'] + 1} ({tbl['num_rows']}×{tbl['num_cols']}):\n{rows_str}")

    prompt = f"""You are reviewing TABLES in a technical document. Check each table for:
1. Missing headers or unclear column names
2. Empty cells that should have values
3. Inconsistent units across rows
4. Test results without pass/fail criteria, actual measurements, or judgments
5. Numerical values without units
6. Calculation errors or suspicious values
7. Consistent formatting across rows

## Tables:
{chr(10).join(tables_text)}

## Output Format:
Return a JSON array of findings. Each finding must be:
{{
  "category": "CATEGORY_ID",
  "severity": "CRITICAL|MAJOR|MINOR|SUGGESTION",
  "page": "-",
  "section": "Table X",
  "comment": "detailed description"
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
            if severity not in SEVERITY_LEVELS:
                severity = "MINOR"

            valid_findings.append({
                "category": category,
                "severity": severity,
                "page": str(f.get("page", "-")),
                "section": str(f.get("section", "-")),
                "comment": str(f.get("comment", "")),
                "source": source,
            })

        return valid_findings
    except json.JSONDecodeError:
        return []


def _deduplicate_findings(findings):
    """Remove duplicate or very similar findings."""
    seen = set()
    unique = []
    for f in findings:
        # Create a rough key from category + first 100 chars of comment
        key = f"{f['category']}:{f['comment'][:100].lower()}"
        if key not in seen:
            seen.add(key)
            unique.append(f)
    return unique
