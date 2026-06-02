"""
Review Engine Module — Commercial Grade v2.0
Multi-pass LLM review + Local Automated Checks for engineering documents.
Detects 14 error pattern categories based on real TICO reviewer feedback.
"""

import os
import gc
import json
import re
import math
import difflib
import hashlib
from datetime import datetime
from collections import Counter, defaultdict
from ollama import Client


# ============================================================
# REVIEW CATEGORIES (expanded from 12 → 16 based on real TICO reviews)
# ============================================================
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
    # --- NEW CATEGORIES from real TICO reviews ---
    "DECIMAL_CONSISTENCY": {
        "name": "Decimal Digit Consistency",
        "icon": "🔢",
        "description": "Inconsistent number of decimal places within table columns or value groups",
    },
    "TABLE_QUALITY": {
        "name": "Table Quality",
        "icon": "📊",
        "description": "Duplicate tables, missing legends, unclear column headers, broken table formatting",
    },
    "SUBSCRIPT_FORMATTING": {
        "name": "Subscript/Superscript Errors",
        "icon": "⬇️",
        "description": "Broken subscripts (<sub> tags in text), caret (^) instead of superscript, formatting artifacts",
    },
    "DATASHEET_COPY_ERROR": {
        "name": "Datasheet Copy Error",
        "icon": "📋",
        "description": "Orphan references from datasheet copy-paste, incomplete notes, wrong figure/equation numbers from datasheets",
    },
    "TOC_VALIDATION": {
        "name": "TOC & Heading Structure",
        "icon": "📚",
        "description": "Table of contents mismatches, broken section numbering, and heading hierarchy problems",
    },
    # --- Standards compliance categories (F3) ---
    "ISO26262_COMPLIANCE": {"name": "ISO 26262 Compliance", "icon": "🛡️",
        "description": "Required ISO 26262 functional-safety elements (safety goals, ASIL, HARA, FMEA)"},
    "IEC61508_COMPLIANCE": {"name": "IEC 61508 Compliance", "icon": "⚙️",
        "description": "Required IEC 61508 elements (SIL, safety function, PFD/PFH, proof test, SFF, HFT)"},
    "AUTOSAR_NAMING": {"name": "AUTOSAR Naming", "icon": "🔤",
        "description": "AUTOSAR component/port/interface naming-convention violations"},
    "FMEA_COVERAGE": {"name": "FMEA Coverage", "icon": "🧯",
        "description": "FMEA table completeness — cause/effect/detection/mitigation columns present and populated"},
    "TRACEABILITY": {"name": "Traceability", "icon": "🔗",
        "description": "Requirement → design → test traceability gaps"},
}

SEVERITY_LEVELS = {
    "CRITICAL": {"label": "Critical", "color": "#ff4444", "weight": 4},
    "MAJOR": {"label": "Major", "color": "#ff8800", "weight": 3},
    "MINOR": {"label": "Minor", "color": "#ffcc00", "weight": 2},
}


# ============================================================
# DETERMINISM & STRUCTURED-OUTPUT CONFIG
# A fixed seed + temperature 0 + a JSON schema make the same
# (document, model) pair produce the same report every run.
# Override the seed via env if a deployment needs to.
# ============================================================
LLM_SEED = int(os.environ.get("DOCAI_LLM_SEED", "42"))

# JSON schema handed to Ollama's `format` parameter so the model returns a
# strictly-parseable findings array. `evidence` is REQUIRED — it forces the
# model to ground every finding in an exact quote, which we then verify.
FINDINGS_JSON_SCHEMA = {
    "type": "array",
    "items": {
        "type": "object",
        "properties": {
            "category": {"type": "string"},
            "severity": {"type": "string", "enum": ["CRITICAL", "MAJOR", "MINOR"]},
            "page": {"type": "string"},
            "section": {"type": "string"},
            "comment": {"type": "string"},
            "fix": {"type": "string"},
            "fix_type": {"type": "string", "enum": ["AUTO", "MANUAL"]},
            "evidence": {"type": "string"},
        },
        "required": ["category", "severity", "comment", "evidence"],
    },
}


# Schema for the self-verification (critic) pass — one verdict per finding.
VERDICTS_JSON_SCHEMA = {
    "type": "array",
    "items": {
        "type": "object",
        "properties": {
            "id": {"type": "integer"},
            "keep": {"type": "boolean"},
            "confidence": {"type": "number"},
            "reason": {"type": "string"},
        },
        "required": ["id", "keep"],
    },
}


def _llm_options(num_predict=4096, seed=None):
    """Standard deterministic options for every findings-generating chat call."""
    return {"temperature": 0, "seed": seed if seed is not None else LLM_SEED, "num_predict": num_predict}


def _chat_findings(client, model, prompt, source, images=None, num_predict=4096, seed=None):
    """
    Single entry point for every findings-producing LLM call.

    Centralizes determinism (seed + temperature 0), structured output
    (`format=FINDINGS_JSON_SCHEMA`), retry on quota errors, and parsing.
    Raises on failure so the caller can record a per-pass error.
    """
    message = {"role": "user", "content": prompt}
    if images:
        message["images"] = images
    response = _chat_with_retry(
        client.chat,
        model=model,
        messages=[message],
        format=FINDINGS_JSON_SCHEMA,
        options=_llm_options(num_predict, seed=seed),
    )
    reply = response["message"]["content"] if isinstance(response, dict) else response.message.content
    return _parse_llm_findings(reply, source)


# ============================================================
# OLLAMA CLIENT
# ============================================================
VISION_MODEL_HINTS = (
    "vl", "vision", "llava", "bakllava", "minicpm-v", "moondream",
)


def is_vision_model(name):
    """Heuristic: does this Ollama model accept images? Used to gate the vision pass."""
    n = (name or "").lower()
    return any(h in n for h in VISION_MODEL_HINTS)


def load_glossary(path=None):
    """
    Load an optional per-project glossary/rules file (JSON). Replaces hardcoded,
    document-specific rules with config the user controls. Path comes from the
    arg or the DOCAI_GLOSSARY env var. Returns {} when absent/invalid.

    Expected shape (all keys optional):
      {
        "acronyms": {"HVDCDC": "High-Voltage DC-DC converter"},
        "canonical_terms": ["Use 'Maximum Ratings' (not 'Max Ratings')"],
        "known_issues": ["A device rating equal to its requirement has no margin"],
        "notes": "free-form project guidance"
      }
    """
    path = path or os.environ.get("DOCAI_GLOSSARY")
    if not path or not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception as e:
        print(f"[glossary] failed to load {path}: {e}")
        return {}


def format_glossary(glossary):
    """Render a glossary dict into a prompt block. Empty string when nothing useful."""
    if not glossary:
        return ""
    lines = ["## PROJECT GLOSSARY & RULES (treat as authoritative)"]
    acronyms = glossary.get("acronyms") or {}
    if isinstance(acronyms, dict) and acronyms:
        lines.append("Acronyms / canonical expansions:")
        lines += [f"- {k}: {v}" for k, v in list(acronyms.items())[:80]]
    for key, header in (("canonical_terms", "Preferred terminology"),
                        ("known_issues", "Known issue patterns to watch for")):
        items = glossary.get(key) or []
        if isinstance(items, list) and items:
            lines.append(f"{header}:")
            lines += [f"- {str(it)}" for it in items[:80]]
    notes = glossary.get("notes")
    if notes:
        lines.append(f"Notes: {notes}")
    return "\n".join(lines) if len(lines) > 1 else ""


def _is_quota_error(e):
    return any(k in str(e).lower() for k in ["429", "rate", "quota", "limit", "upgrade"])


def _chat_with_retry(client_chat_fn, max_retries=3, **kwargs):
    """
    Call a client.chat() fn with exponential backoff on quota/rate errors.
    Raises only after all retries are exhausted so transient 429s don't kill a pass.
    """
    import time
    last_err = None
    for attempt in range(max_retries):
        try:
            return client_chat_fn(**kwargs)
        except Exception as e:
            last_err = e
            if _is_quota_error(e):
                wait = 2 ** attempt  # 1s, 2s, 4s
                print(f"[retry] quota/rate error (attempt {attempt+1}/{max_retries}), waiting {wait}s: {str(e)[:80]}")
                time.sleep(wait)
            else:
                raise  # not a quota issue — fail immediately
    raise last_err


def create_ollama_client(api_key, host="https://ollama.com"):
    """
    Create an LLM client for the given host.

    - OpenAI-compatible hosts (FreeLLMAPI, LM Studio, vLLM, any /v1 endpoint) get
      an OpenAICompatClient adapter that speaks the same .chat()/.list() interface.
    - Otherwise a native Ollama Client. A Bearer token is sent only when an API key
      is given, so a LOCAL Ollama (http://localhost:11434, no key) works the same as cloud.

    Name kept as create_ollama_client for backward compatibility across the codebase.
    """
    try:
        from llm_client import is_openai_compat_host, OpenAICompatClient
        if is_openai_compat_host(host):
            return OpenAICompatClient(api_key, host)
    except ImportError:
        pass
    headers = {"Authorization": f"Bearer {api_key}"} if api_key else {}
    return Client(host=host, headers=headers)


class FailoverOllamaClient:
    """
    A wrapper around Ollama Client that holds multiple API keys and
    automatically rotates to the next key when the current one fails.
    """

    def __init__(self, api_keys, host="https://ollama.com"):
        if isinstance(api_keys, str):
            api_keys = [api_keys]
        self.api_keys = [k.strip() for k in api_keys if k.strip()]
        self.host = host
        self.current_index = 0
        self._clients = [create_ollama_client(k, host) for k in self.api_keys]

    def _get_client(self):
        return self._clients[self.current_index]

    def _rotate(self):
        """Move to the next API key. Returns True if a new key is available."""
        old = self.current_index
        self.current_index = (self.current_index + 1) % len(self._clients)
        return self.current_index != old  # True if we actually moved to a different key

    def chat(self, **kwargs):
        """Call chat with automatic failover across all API keys."""
        last_error = None
        for _ in range(len(self._clients)):
            try:
                return self._get_client().chat(**kwargs)
            except Exception as e:
                last_error = e
                err_str = str(e).lower()
                # Only failover on quota/rate/auth errors, not on model errors
                if any(kw in err_str for kw in ["rate", "limit", "quota", "unauthorized", "forbidden", "429", "503"]):
                    print(f"[Failover] Key #{self.current_index + 1} failed ({str(e)[:80]}), rotating...")
                    if not self._rotate():
                        break  # Only 1 key, can't rotate
                else:
                    raise e  # It's not a key issue, re-raise immediately
        raise last_error  # All keys exhausted

    def list(self):
        """List models using the current key."""
        return self._get_client().list()


HARDCODED_POOL_KEYS = [
    "b02f4829385f408da3696bc5e1792cf6.dW2QqbUHt8qGV-xGFOXnJx92",
    "6bec99a28c5d45e6904471700d8a1981.tSYIbCUyvtFRbxfCxuv3J4ll",
    "7469dd11682643e0897e94d964fa5ddd.eKsnbbhTmTrYpd_RaVkPkczC",
    "2980579cd2c94ccfbafaf13c8b8cd2a2.PtzhKu_WFDX3kWAoTYJ2x5mo",
    "1ad8ee352f525427eb9dff1921c4167b2.t0UvepQrOx5bHSZPpy7eqke",
    "28a9f40eb3442caaea1d72bd31d484a.-BioB55Cjx7GQ3xh-3G2knQC",
    "e2eee2db379c4d4cb4633139c7e1f5f8.v8DzylMngpl9BVc-UQlTso-8",
    "04959e75c10047c7852574c9f31bd21.WtQ9V5MH3hqxm2BD3l3dp3Vl",
    "0a430ede41f54938bcd8ebbbfc0ac70.wEC7y9zMDSFUQVNs47j5aqWX",
    "eb326f38a4f045e69aee164ea0052d98.7KTdjcA0L7VNyiQvbTTMYwLL",
    "d07f97c51baa4fb7936bb33b0071e368.TEKhpo-yT1_CoP1oLGhy5RMK",
    "186f70c3ea604c509420b2ae939d5c3.-Tql3V7lzrrN9rM41QWqGNaF",
    "a1d3960a079c4cbe9c87fd0f0d0a49.IFC0dMP9pVjTL6z7LS5xuVr2",
    "33cc1ef927d24c88ba3d981c86c918df.uqh6geLCsH3rJOrEX3zKd5TF"
]

# Optional: supply/rotate pool keys via env (comma-separated) WITHOUT editing code.
# Env keys are tried first, then the hardcoded list. Set OLLAMA_POOL_KEYS in Render
# or backend/.env to move these out of source control later.
_env_pool_keys = [k.strip() for k in os.environ.get("OLLAMA_POOL_KEYS", "").split(",") if k.strip()]
if _env_pool_keys:
    HARDCODED_POOL_KEYS = _env_pool_keys + HARDCODED_POOL_KEYS


def create_failover_client(api_keys, host="https://ollama.com"):
    """Create a FailoverOllamaClient from a list of API keys (or a single key)."""
    return FailoverOllamaClient(api_keys, host)


def test_connection(api_key, host="https://ollama.com"):
    """
    Test connection and return available models.

    api_key may be a comma-separated list (failover keys). Model listing is a
    cheap GET that doesn't burn generation quota, so we try each key in turn and
    return on the first that succeeds — never send the whole list as one bearer
    token (that's invalid and is why a multi-key setup failed the connection test).
    """
    keys = [k.strip() for k in (api_key or "").split(",") if k.strip()]
    # Fallback to hardcoded pool keys if none given on the cloud host
    if not keys and (host or "").lower() == "https://ollama.com":
        keys = [k for k in (globals().get("HARDCODED_POOL_KEYS") or []) if k]
    # Local hosts need no key — try once with an empty key.
    if not keys:
        keys = [""]

    last_error = None
    for key in keys:
        try:
            client = create_ollama_client(key, host)
            models = client.list()
            model_names = []
            if hasattr(models, "models"):
                model_names = [m.model for m in models.models]
            elif isinstance(models, dict) and "models" in models:
                model_names = [m.get("name", m.get("model", "unknown")) for m in models["models"]]
            return {
                "success": True,
                "models": model_names,
                "vision_models": [m for m in model_names if is_vision_model(m)],
            }
        except Exception as e:
            last_error = e
            continue
    return {"success": False, "error": str(last_error) if last_error else "Could not connect"}


# ============================================================
# MAIN REVIEW ORCHESTRATOR
# ============================================================
def review_document(client, model, parsed_doc, progress_callback=None, review_mode="pro", vision_model=None, status_out=None, standards=None, glossary=None):
    """
    Perform comprehensive multi-pass review of a parsed document.
    Uses 'model' for text/table review and 'vision_model' for image review.

    Review Pipeline:
      1. Local Automated Checks (100% consistent, no LLM)
      2. Pass A — Text Quality (grammar, spelling, terminology per chunk)
      3. Pass B — Technical Accuracy (units, calculations, orphan refs per chunk)
      4. Table-specific deep review
      5. Full-document consistency check
      6. Image review with vision model

    If `status_out` (a dict) is provided, it is populated with per-pass status so
    the caller can tell the user when AI passes failed instead of silently
    returning a local-only report. Shape:
        {"passes": [{"name", "ok", "error", "count"}], "failed_count": int}
    """
    findings = []
    text_model = model
    img_model = vision_model or model
    pass_status = []  # [{name, ok, error, count}]

    def _record(name, ok, count=0, error=None):
        pass_status.append({"name": name, "ok": ok, "count": count, "error": error})
    
    # Define which categories to use based on mode
    if review_mode == "normal":
        active_categories = ["GRAMMAR_SPELLING", "TERMINOLOGY_CONSISTENCY", "UNITS_CALCULATIONS"]
    else:
        active_categories = list(REVIEW_CATEGORIES.keys())

    try:
        # ── STEP 1: Local Automated Checks (no LLM, 100% consistent) ──
        if progress_callback:
            progress_callback("Running automated quality checks...", 8)
        
        local_findings = _run_local_checks(parsed_doc)
        if local_findings:
            findings.extend(local_findings)

        # ── STEP 1b: Standards compliance checks (deterministic, opt-in) ──
        if standards:
            try:
                from standards.checker import run_standards_checks
                std_findings, std_checklist = run_standards_checks(parsed_doc, standards)
                if std_findings:
                    findings.extend(std_findings)
                if status_out is not None:
                    status_out["compliance_checklist"] = std_checklist
                _record("standards", True, count=len(std_findings))
            except Exception as e:
                _record("standards", False, error=str(e))

        if progress_callback:
            progress_callback(f"Found {len(local_findings)} issues from automated checks. Starting AI analysis...", 12)

        # ── STEP 2: LLM-powered multi-pass review ──
        from doc_parser import get_document_summary, get_section_chunks_smart

        doc_summary = get_document_summary(parsed_doc)
        # Per-review glossary (from the user/UI) overrides the env-file glossary.
        glossary_text = format_glossary(glossary if glossary else load_glossary())
        chunks = get_section_chunks_smart(parsed_doc, target_chars=8000)
        total_chunks = len(chunks)

        # Pass A — Text Quality + Technical combined (per chunk).
        # SPEED: chunks are independent LLM calls, so run them concurrently. With a
        # multi-key failover pool this also spreads load across keys. Tune with
        # DOCAI_CONCURRENCY (default 5). Set to 1 to force sequential.
        import concurrent.futures as _cf
        concurrency = max(1, int(os.environ.get("DOCAI_CONCURRENCY", "5")))
        seeds = [LLM_SEED, 137] if review_mode == "max" else [LLM_SEED]

        def _do_chunk(args):
            idx, chunk = args
            out = []
            for sd in seeds:
                out.extend(_review_chunk_multipass(
                    client, text_model, chunk, doc_summary, idx + 1,
                    active_categories, glossary_text, seed=sd))
            return out

        chunk_failures = 0
        chunk_count = 0
        last_chunk_error = None
        done = 0
        if concurrency == 1 or total_chunks <= 1:
            iterator = ((i, c) for i, c in enumerate(chunks))
            for args in iterator:
                try:
                    findings.extend(_do_chunk(args))
                    chunk_count += 1
                except Exception as e:
                    chunk_failures += 1
                    last_chunk_error = str(e)
                done += 1
                if progress_callback:
                    progress_callback(f"AI Pass: chunk {done}/{total_chunks}...", 15 + int((done / max(total_chunks, 1)) * 50))
        else:
            with _cf.ThreadPoolExecutor(max_workers=concurrency) as pool:
                futures = {pool.submit(_do_chunk, (i, c)): i for i, c in enumerate(chunks)}
                for fut in _cf.as_completed(futures):
                    try:
                        findings.extend(fut.result())
                        chunk_count += 1
                    except Exception as e:
                        chunk_failures += 1
                        last_chunk_error = str(e)
                    done += 1
                    if progress_callback:
                        progress_callback(f"AI Pass: chunk {done}/{total_chunks}...", 15 + int((done / max(total_chunks, 1)) * 50))
        gc.collect()

        # A pass "fails" only if EVERY chunk failed (e.g. bad model / no quota).
        if total_chunks and chunk_failures == total_chunks:
            _record("text_chunks", False, count=0, error=last_chunk_error)
        else:
            _record("text_chunks", True, count=chunk_count,
                    error=(f"{chunk_failures} of {total_chunks} chunks failed" if chunk_failures else None))

        # ── STEP 3: Full-document cross-reference and consistency check ──
        if progress_callback:
            progress_callback("Checking cross-document consistency & terminology...", 68)
        
        if review_mode in ("pro", "max"):
            try:
                consistency_findings = _review_consistency_with_llm(client, text_model, doc_summary, glossary_text)
                if consistency_findings:
                    findings.extend(consistency_findings)
                _record("consistency", True, count=len(consistency_findings))
            except Exception as e:
                _record("consistency", False, error=str(e))

        # ── B2: Category-specialized deep passes (max mode) ──
        # Focused prompts (units-only, terminology-only) get higher recall per
        # domain than the general chunk pass. Gated to max while validating.
        if review_mode == "max":
            if progress_callback:
                progress_callback("Max: focused units & calculations pass...", 70)
            try:
                units_findings = _review_units_pass(client, text_model, parsed_doc, glossary_text)
                if units_findings:
                    findings.extend(units_findings)
                _record("units_pass", True, count=len(units_findings))
            except Exception as e:
                _record("units_pass", False, error=str(e))

            if progress_callback:
                progress_callback("Max: focused terminology & naming pass...", 72)
            try:
                term_findings = _review_terminology_pass(client, text_model, parsed_doc, glossary_text)
                if term_findings:
                    findings.extend(term_findings)
                _record("terminology_pass", True, count=len(term_findings))
            except Exception as e:
                _record("terminology_pass", False, error=str(e))

        # ── STEP 4: Table-specific review (LLM) ──
        if parsed_doc.get("tables"):
            if progress_callback:
                progress_callback("Deep-reviewing tables and data...", 75)

            try:
                if review_mode in ("pro", "max"):
                    table_findings = _review_tables_with_llm(client, text_model, parsed_doc)
                elif "UNITS_CALCULATIONS" in active_categories:
                    table_findings = _review_tables_with_llm(client, text_model, parsed_doc, ["UNITS_CALCULATIONS"])
                else:
                    table_findings = []
                if table_findings:
                    findings.extend(table_findings)
                _record("tables", True, count=len(table_findings))
            except Exception as e:
                _record("tables", False, error=str(e))

        # ── STEP 5: Image-specific review (use vision model) ──
        if parsed_doc.get("images") and is_vision_model(img_model):
            total_images = len([i for i in parsed_doc["images"] if i.get("full_b64") and not i.get("is_small")])
            if progress_callback:
                progress_callback(f"Reviewing {total_images} images/diagrams with {img_model}...", 82)

            if review_mode in ("pro", "max"):
                image_errors = []
                try:
                    image_findings = _review_images_with_llm(client, img_model, parsed_doc, doc_summary, progress_callback, errors=image_errors)
                    if image_findings:
                        findings.extend(image_findings)
                    # The pass only fully fails if every image errored and none succeeded.
                    if image_errors and not image_findings:
                        _record("images", False, error="; ".join(image_errors[:3]))
                    else:
                        _record("images", True, count=len(image_findings),
                                error=("; ".join(image_errors[:3]) if image_errors else None))
                except Exception as e:
                    _record("images", False, error=str(e))

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
    if progress_callback:
        progress_callback("Deduplicating and finalizing findings...", 92)
    
    if findings:
        # Drop hallucinated LLM findings + anchor survivors to a paragraph (0.4/0.8).
        findings = _ground_and_anchor_findings(findings, parsed_doc)
        # Critic pass (Pro/Max): second opinion that SCORES confidence (no longer
        # deletes — recall-first). Local findings stay at 1.0.
        if review_mode in ("pro", "max"):
            try:
                findings = _critic_filter_findings(client, text_model, findings, progress_callback)
            except Exception as e:
                print(f"Critic pass skipped: {e}")
        # Cap ungrounded LLM findings BEFORE the Max precision cut so they're the
        # ones trimmed (their evidence isn't verbatim in the document).
        for f in findings:
            if f.get("grounded") is False:
                f["confidence"] = min(float(f.get("confidence", 0.5) or 0.5), 0.4)
        # Max = strictest: keep only high-confidence findings (precision over recall).
        # Pro/Normal keep everything — use the confidence slider to trim.
        if review_mode == "max":
            findings = [f for f in findings if float(f.get("confidence", 1.0)) >= 0.5]
        findings = _deduplicate_findings(findings)
        # Sort by severity, then confidence — most trustworthy issues first.
        findings.sort(
            key=lambda f: (
                SEVERITY_LEVELS.get(f.get("severity", "MINOR"), {}).get("weight", 0),
                float(f.get("confidence", 0) or 0),
            ),
            reverse=True,
        )

    # Number findings and assign fix_type
    for idx, f in enumerate(findings, 1):
        f["id"] = idx
        if "fix_type" not in f:
            f["fix_type"] = _classify_fix_type(f)
        if "status" not in f:
            f["status"] = "OPEN"
        f.setdefault("confidence", 1.0 if not str(f.get("source", "")).startswith(("llm", "vision")) else 0.7)
        # Ungrounded LLM findings (evidence not verbatim in the doc) are kept but
        # capped at lower confidence so the slider/Max mode can trim them.
        if f.get("grounded") is False:
            f["confidence"] = min(float(f.get("confidence", 0.5)), 0.4)

    # Publish per-pass status so the caller can warn the user about failed AI passes.
    if status_out is not None:
        failed = [p for p in pass_status if not p["ok"]]
        status_out["passes"] = pass_status
        status_out["failed_count"] = len(failed)
        if failed:
            # Give a more specific message when it is a quota error
            all_errors = " ".join(p.get("error", "") for p in failed)
            if "429" in all_errors or "quota" in all_errors.lower() or "upgrade" in all_errors.lower():
                msg = (
                    f"{len(failed)} AI pass(es) hit the Ollama Cloud quota limit "
                    f"({', '.join(p['name'] for p in failed)}). "
                    f"To fix: add a second API key in Settings (paste both keys separated by a comma) "
                    f"so the tool auto-rotates when one key is exhausted. "
                    f"Or switch to a local Ollama for unlimited usage."
                )
            else:
                msg = (
                    f"{len(failed)} AI pass(es) failed ({', '.join(p['name'] for p in failed)}). "
                    f"The report may be incomplete — check the model name and API key/host."
                )
            status_out["warning"] = msg

    return findings


def _classify_fix_type(finding):
    """Classify whether a finding can be auto-fixed or needs manual intervention."""
    cat = finding.get("category", "")
    comment = finding.get("comment", "").lower()
    
    # Auto-fixable: spelling errors, simple terminology swaps
    if cat == "GRAMMAR_SPELLING":
        # Only if the fix mentions a clear replacement
        fix = finding.get("fix", "").lower()
        if any(kw in fix for kw in ["change", "replace", "should be", "correct to"]):
            return "AUTO"
    
    # Everything else is manual
    return "MANUAL"


# ============================================================
# LOCAL AUTOMATED CHECKS (No LLM, 100% Consistent)
# ============================================================
# Fonts that are legitimately different from the body font (symbols, math, icons).
# Flagging these as "font inconsistency" is a pure false positive.
SYMBOL_FONTS = {
    "symbol", "wingdings", "wingdings 2", "wingdings 3", "webdings",
    "cambria math", "mt extra", "marlett", "segoe mdl2 assets",
}

# Common power-rail / net names that look like "value+unit" but are identifiers,
# not measurements — do not flag these for missing unit spacing.
COMMON_RAIL_NAMES = {
    "5V", "3V3", "3V0", "5V0", "1V8", "1V2", "1V0", "12V", "24V", "48V", "0V",
}


def _run_local_checks(parsed_doc):
    """
    Run all local checks that don't need LLM.
    These produce identical results every run — 100% consistent.
    Note: TOC/heading checks are Word-only; Excel docs have no toc/headings so they skip silently.
    """
    findings = []

    # 1. Font & formatting checks
    findings.extend(_check_font_consistency(parsed_doc))

    # 2. Decimal digit consistency in tables
    findings.extend(_check_decimal_consistency(parsed_doc))

    # 3. Cross-reference validation
    findings.extend(_check_cross_references(parsed_doc))

    # 4. Table duplication detection
    findings.extend(_check_table_duplication(parsed_doc))

    # 5. Subscript/formatting errors
    findings.extend(_check_subscript_errors(parsed_doc))

    # 6. Orphan datasheet references
    findings.extend(_check_orphan_references(parsed_doc))

    # 7. TOC ↔ Heading sync (Word-only)
    findings.extend(_check_toc_heading_sync(parsed_doc))

    # 8. Section number continuity (Word-only)
    findings.extend(_check_section_number_continuity(parsed_doc))

    # 9. Heading hierarchy enforcement (Word-only)
    findings.extend(_check_heading_hierarchy(parsed_doc))

    # 10. Spacing and tab errors
    findings.extend(_check_spacing_errors(parsed_doc))

    # 11. Repeated words
    findings.extend(_check_repeated_words(parsed_doc))

    # 12. Unmatched brackets
    findings.extend(_check_unmatched_brackets(parsed_doc))

    # 13. Empty paragraph pagination
    findings.extend(_check_empty_paragraphs(parsed_doc))

    # 14. Min/Typ/Max logic validator
    findings.extend(_check_min_typ_max_tables(parsed_doc))

    # 15. Engineering standard unit casing/spacing
    findings.extend(_check_unit_standardization(parsed_doc))

    return findings


def _check_font_consistency(parsed_doc):
    """Check for font name and size inconsistencies."""
    findings = []
    fmt = parsed_doc.get("formatting", {})

    # Check for font inconsistencies
    default_font = fmt.get("default_font")
    if default_font:
        font_mismatches = set()
        for section in parsed_doc["sections"]:
            for para in section["paragraphs"]:
                for run in para.get("runs", []):
                    if "font" in run and run["font"] != default_font and run["text"].strip():
                        # Skip symbol/math/icon fonts — these are legitimately different.
                        if run["font"].strip().lower() in SYMBOL_FONTS:
                            continue
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
                "fix_type": "MANUAL",
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
                "fix_type": "MANUAL",
            })

    return findings


def _check_decimal_consistency(parsed_doc):
    """
    Check if decimal places are consistent within table columns.
    Pattern 1 from real TICO reviews — VERY HIGH frequency.
    """
    findings = []
    tables = parsed_doc.get("tables", [])
    
    number_pattern = re.compile(r'^-?\d+\.\d+$')
    
    for tbl in tables:
        tbl_name = tbl.get("name", f"Table {tbl.get('index', 0) + 1}")
        rows = tbl.get("rows", [])
        if len(rows) < 3:  # Need at least header + 2 data rows
            continue
        
        num_cols = tbl.get("num_cols", 0)
        
        for col_idx in range(num_cols):
            decimal_counts = {}  # decimal_places -> list of values
            for row_idx, row in enumerate(rows[1:], 1):  # Skip header
                if col_idx >= len(row):
                    continue
                cell = row[col_idx].strip()
                
                # Extract numbers from cells
                numbers_in_cell = number_pattern.findall(cell)
                if not numbers_in_cell:
                    # Try to find numbers with units like "3.3V" or "3.300 V"
                    nums = re.findall(r'(-?\d+\.\d+)\s*[a-zA-Z°Ω%]*', cell)
                    numbers_in_cell = nums
                
                for num_str in numbers_in_cell:
                    if '.' in num_str:
                        decimal_places = len(num_str.split('.')[-1])
                        if decimal_places not in decimal_counts:
                            decimal_counts[decimal_places] = []
                        decimal_counts[decimal_places].append(f"Row {row_idx}: {cell[:40]}")
            
            # Flag if column has numbers with different decimal places
            if len(decimal_counts) > 1 and sum(len(v) for v in decimal_counts.values()) >= 3:
                # Find the most common decimal count
                most_common = max(decimal_counts, key=lambda k: len(decimal_counts[k]))
                problem_counts = {k: v for k, v in decimal_counts.items() if k != most_common}
                
                if problem_counts:
                    problem_details = []
                    for dc, examples in problem_counts.items():
                        problem_details.append(f"{dc} decimal places ({len(examples)} values, e.g., {examples[0][:30]})")
                    
                    col_header = rows[0][col_idx].strip() if col_idx < len(rows[0]) else f"Column {col_idx+1}"
                    findings.append({
                        "category": "DECIMAL_CONSISTENCY",
                        "severity": "MAJOR",
                        "page": "-",
                        "section": tbl_name,
                        "comment": (
                            f"Inconsistent decimal places in column '{col_header}'. "
                            f"Most values use {most_common} decimal places, but found: "
                            f"{'; '.join(problem_details)}. "
                            f"All values in this column should use {most_common} decimal places."
                        ),
                        "fix": f"Standardize all values in column '{col_header}' to {most_common} decimal places.",
                        "source": "local",
                        "fix_type": "MANUAL",
                    })
    
    return findings


def _check_cross_references(parsed_doc):
    """
    Verify that Figure X, Table X, Equation X, Section X references point to actual items.
    Pattern 2 from real TICO reviews — VERY HIGH frequency.
    """
    findings = []
    raw_text = parsed_doc.get("raw_text", "")
    
    # Build index of actual figures/tables/sections
    actual_figures = set()
    actual_tables = set()
    actual_equations = set()
    actual_sections = set()
    
    # From sections headings
    for section in parsed_doc.get("sections", []):
        heading = section.get("heading", "")
        # Check for figure mentions in headings
        fig_match = re.findall(r'Figure\s+(\d+[-.]?\d*)', heading, re.IGNORECASE)
        actual_figures.update(fig_match)
        
        tbl_match = re.findall(r'Table\s+(\d+[-.]?\d*)', heading, re.IGNORECASE)
        actual_tables.update(tbl_match)
        
        eq_match = re.findall(r'Equation\s+(\d+[-.]?\d*)', heading, re.IGNORECASE)
        actual_equations.update(eq_match)
        
        # Track section numbers
        sec_match = re.match(r'^(\d+(?:\.\d+)*)', heading.strip())
        if sec_match:
            actual_sections.add(sec_match.group(1))
    
    # From table names
    for tbl in parsed_doc.get("tables", []):
        tbl_name = tbl.get("name", "")
        num_match = re.findall(r'Table\s+(\d+[-.]?\d*)', tbl_name, re.IGNORECASE)
        actual_tables.update(num_match)
    
    # Scan paragraphs for captions. Word captions use the "Caption" paragraph
    # style and SEQ fields (which python-docx renders as "Figure 5" / "Table 3"),
    # often WITHOUT a trailing ':' or '—'. So we treat a paragraph as a caption
    # definition when EITHER its style is a caption style OR the label appears at
    # the very start of the paragraph (the canonical caption position).
    caption_lead = re.compile(r'^\s*(Figure|Table|Equation)\s+(\d+[-.]?\d*)', re.IGNORECASE)
    caption_sep = re.compile(r'(Figure|Table|Equation)\s+(\d+[-.]?\d*)\s*[:\-—]', re.IGNORECASE)

    def _record_label(kind, num):
        kind = kind.lower()
        if kind == "figure":
            actual_figures.add(num)
        elif kind == "table":
            actual_tables.add(num)
        elif kind == "equation":
            actual_equations.add(num)

    for section in parsed_doc.get("sections", []):
        for para in section.get("paragraphs", []):
            text = para.get("text", "")
            if not text:
                continue
            is_caption_style = "caption" in str(para.get("style", "")).lower()

            # Definition with a separator: "Figure 5: ..." anywhere in the line.
            for kind, num in caption_sep.findall(text):
                _record_label(kind, num)

            if is_caption_style:
                # The whole paragraph is a caption — count every label in it.
                for kind, num in re.findall(r'(Figure|Table|Equation)\s+(\d+[-.]?\d*)', text, re.IGNORECASE):
                    _record_label(kind, num)
            else:
                # A label at the very start of a normal paragraph is the
                # canonical caption position ("Figure 72 Pin output diagram").
                lead = caption_lead.match(text)
                if lead:
                    _record_label(lead.group(1), lead.group(2))
    
    # Now scan for all references in body text
    ref_pattern = re.compile(r'(?:see\s+|refer\s+to\s+|in\s+|from\s+)?(Figure|Table|Equation|Section)\s+(\d+[-.]?\d*)', re.IGNORECASE)
    
    issues_found = defaultdict(list)
    
    for section in parsed_doc.get("sections", []):
        section_name = section.get("heading", "Unknown")
        for para in section.get("paragraphs", []):
            text = para.get("text", "")
            page = str(para.get("page", "-"))
            
            for match in ref_pattern.finditer(text):
                ref_type = match.group(1).capitalize()
                ref_num = match.group(2)
                
                # Skip if it's a definition/caption itself
                context = text[max(0, match.start()-5):match.end()+10]
                if re.search(r'[:\-—]', context[len(match.group(0)):]):
                    continue
                
                is_valid = False
                if ref_type == "Figure":
                    is_valid = ref_num in actual_figures
                elif ref_type == "Table":
                    is_valid = ref_num in actual_tables
                elif ref_type == "Equation":
                    is_valid = ref_num in actual_equations
                elif ref_type == "Section":
                    is_valid = ref_num in actual_sections
                
                if not is_valid:
                    key = f"{ref_type} {ref_num}"
                    if key not in issues_found:
                        issues_found[key] = {
                            "page": page,
                            "section": section_name,
                            "ref_type": ref_type,
                            "ref_num": ref_num,
                        }
    
    # PRECISION GUARD (Phase 0): our caption/figure/table extraction only
    # recognizes a subset of caption styles, so the "actual" sets can be
    # incomplete. If we would flag MORE broken references of a type than we
    # actually detected of that type, our index — not the document — is the
    # problem, so suppress that type to avoid a flood of false positives.
    # (Phase 1.2 replaces this with proper SEQ-field / Caption detection.)
    actual_counts = {
        "Figure": len(actual_figures),
        "Table": len(actual_tables),
        "Equation": len(actual_equations),
        "Section": len(actual_sections),
    }
    broken_counts = Counter(info["ref_type"] for info in issues_found.values())
    unreliable_types = {
        rtype for rtype, broken in broken_counts.items()
        if broken > max(actual_counts.get(rtype, 0), 2)
    }

    for key, info in issues_found.items():
        if info["ref_type"] in unreliable_types:
            continue
        findings.append({
            "category": "CROSS_REFERENCE_ACCURACY",
            "severity": "MAJOR",
            "page": info["page"],
            "section": info["section"],
            "comment": (
                f"Broken reference: '{key}' is referenced but does not appear to exist in this document. "
                f"It may be a wrong number, or the referenced item is missing."
            ),
            "fix": f"Verify that {key} exists. If it doesn't, correct the reference to the right number or add the missing item.",
            "evidence": key,
            "source": "local",
            "fix_type": "MANUAL",
        })

    return findings


def _check_table_duplication(parsed_doc):
    """
    Detect near-identical tables that may have been copy-pasted.  
    Pattern 8 from real TICO reviews.
    """
    findings = []
    tables = parsed_doc.get("tables", [])
    
    if len(tables) < 2:
        return findings
    
    # Compute a content hash for each table
    table_hashes = []
    for tbl in tables:
        rows = tbl.get("rows", [])
        # Hash ALL rows — two large tables that differ only after row 50 must not
        # be mistaken for duplicates (and vice versa).
        content = "||".join(
            "|".join(cell.strip().lower() for cell in row)
            for row in rows
        )
        h = hashlib.md5(content.encode()).hexdigest()
        table_hashes.append({
            "hash": h,
            "content": content,
            "name": tbl.get("name", f"Table {tbl.get('index', 0) + 1}"),
            "num_rows": tbl.get("num_rows", 0),
        })
    
    # Compare pairwise
    seen_pairs = set()
    for i in range(len(table_hashes)):
        for j in range(i + 1, len(table_hashes)):
            pair_key = (min(i, j), max(i, j))
            if pair_key in seen_pairs:
                continue
            seen_pairs.add(pair_key)
            
            ti = table_hashes[i]
            tj = table_hashes[j]
            
            # Exact match
            if ti["hash"] == tj["hash"] and ti["content"]:
                findings.append({
                    "category": "TABLE_QUALITY",
                    "severity": "MAJOR",
                    "page": "-",
                    "section": f"{ti['name']} / {tj['name']}",
                    "comment": f"Tables '{ti['name']}' and '{tj['name']}' appear to be exact duplicates ({ti['num_rows']} rows each). This may be a copy-paste error.",
                    "fix": "Remove the duplicate table or differentiate their content.",
                    "source": "local",
                    "fix_type": "MANUAL",
                })
            elif ti["content"] and tj["content"] and len(ti["content"]) > 50:
                # Check similarity
                ratio = difflib.SequenceMatcher(None, ti["content"][:500], tj["content"][:500]).ratio()
                if ratio > 0.85:
                    findings.append({
                        "category": "TABLE_QUALITY",
                        "severity": "MINOR",
                        "page": "-",
                        "section": f"{ti['name']} / {tj['name']}",
                        "comment": f"Tables '{ti['name']}' and '{tj['name']}' are {ratio*100:.0f}% similar. This may be an unintended duplication from the datasheet.",
                        "fix": "Review both tables and remove or merge if they contain essentially the same data.",
                        "source": "local",
                        "fix_type": "MANUAL",
                    })
    
    return findings


def _check_subscript_errors(parsed_doc):
    """
    Detect broken subscripts, HTML tags in text, caret notation in variable names.
    Pattern 4 from real TICO reviews.
    """
    findings = []
    
    # Patterns to detect
    html_tag_pattern = re.compile(r'</?(?:sub|sup|b|i|em|strong)>', re.IGNORECASE)
    caret_pattern = re.compile(r'[A-Z]+\w*\^[\-\d\w]+')  # e.g., VOUT^-0.879
    
    for section in parsed_doc.get("sections", []):
        section_name = section.get("heading", "Unknown")
        for para in section.get("paragraphs", []):
            text = para.get("text", "")
            page = str(para.get("page", "-"))
            
            # Check for HTML tags in text (copy-paste artifact)
            html_matches = html_tag_pattern.findall(text)
            if html_matches:
                findings.append({
                    "category": "SUBSCRIPT_FORMATTING",
                    "severity": "MAJOR",
                    "page": page,
                    "section": section_name,
                    "comment": f"HTML formatting tags found in text: {', '.join(set(html_matches))}. This indicates broken subscript/superscript from copy-paste. Text: '{text[:80]}...'",
                    "fix": "Convert HTML tags to proper Word subscript/superscript formatting (select text → Format → Font → Subscript/Superscript).",
                    "source": "local",
                    "fix_type": "MANUAL",
                })
            
            # Check for caret notation
            caret_matches = caret_pattern.findall(text)
            if caret_matches:
                for cm in caret_matches[:3]:  # Limit reports
                    findings.append({
                        "category": "SUBSCRIPT_FORMATTING",
                        "severity": "MINOR",
                        "page": page,
                        "section": section_name,
                        "comment": f"Caret notation '{cm}' found. Use proper superscript formatting instead of '^'. ",
                        "fix": f"Replace '{cm}' with proper superscript formatting (select the exponent, Format → Font → Superscript).",
                        "source": "local",
                        "fix_type": "MANUAL",
                    })
    
    return findings


def _check_orphan_references(parsed_doc):
    """
    Detect orphan references from datasheet copy-paste.
    Pattern 5 from real TICO reviews — references like "(1)", "(Note 3)", 
    "See Figure 9-4" that are internal to a datasheet but not the document.
    """
    findings = []
    
    # Patterns for datasheet-internal references
    note_ref = re.compile(r'\((?:Note|注)\s*\d+\)')  # (Note 1), (Note 3)
    footnote_ref = re.compile(r'\(\d+\)\s*$')  # (1) at end of cell/line
    datasheet_fig = re.compile(r'(?:See\s+)?Figure\s+\d+-\d+', re.IGNORECASE)  # Figure 9-4 (datasheet style)
    
    for section in parsed_doc.get("sections", []):
        section_name = section.get("heading", "Unknown")
        for para in section.get("paragraphs", []):
            text = para.get("text", "")
            page = str(para.get("page", "-"))
            
            # Check for datasheet note references
            note_matches = note_ref.findall(text)
            for nm in note_matches:
                findings.append({
                    "category": "DATASHEET_COPY_ERROR",
                    "severity": "MAJOR",
                    "page": page,
                    "section": section_name,
                    "comment": f"Orphan datasheet reference '{nm}' found. This note reference likely came from a component datasheet and does not exist in this document.",
                    "fix": f"Remove the orphan reference '{nm}' or add the corresponding note to this document.",
                    "source": "local",
                    "fix_type": "MANUAL",
                })
            
            # Check for datasheet-style figure references (Figure X-Y format)
            ds_fig_matches = datasheet_fig.findall(text)
            for dfm in ds_fig_matches:
                findings.append({
                    "category": "DATASHEET_COPY_ERROR",
                    "severity": "MAJOR",
                    "page": page,
                    "section": section_name,
                    "comment": f"Datasheet-style reference '{dfm}' found. This figure reference uses datasheet numbering (X-Y format) and likely does not exist in this document.",
                    "fix": f"Remove '{dfm}' or replace with the correct figure reference from this document.",
                    "source": "local",
                    "fix_type": "MANUAL",
                })
    
    return findings


# ============================================================
# LLM-POWERED REVIEW FUNCTIONS
# ============================================================
def _cache_key(*parts):
    return hashlib.md5("||".join(str(p) for p in parts).encode("utf-8")).hexdigest()


def _cache_get(key):
    """Read cached findings for a content key. Enabled only when DOCAI_CACHE_DIR is set."""
    cache_dir = os.environ.get("DOCAI_CACHE_DIR")
    if not cache_dir:
        return None
    path = os.path.join(cache_dir, key + ".json")
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return None
    return None


def _cache_put(key, findings):
    cache_dir = os.environ.get("DOCAI_CACHE_DIR")
    if not cache_dir:
        return
    try:
        os.makedirs(cache_dir, exist_ok=True)
        with open(os.path.join(cache_dir, key + ".json"), "w", encoding="utf-8") as fh:
            json.dump(findings, fh)
    except Exception:
        pass


def _review_chunk_multipass(client, model, chunk_text, doc_summary, chunk_num, active_categories, glossary_text="", seed=None):
    """
    Enhanced chunk review with focused, detailed prompt.
    Combines text quality + technical accuracy in one focused pass per chunk.
    """
    cat_list = "\n".join([
        f"- {cid}: {REVIEW_CATEGORIES[cid]['name']} — {REVIEW_CATEGORIES[cid]['description']}"
        for cid in active_categories if cid in REVIEW_CATEGORIES
    ])
    glossary_block = f"\n{glossary_text}\n" if glossary_text else ""

    prompt = f"""You are an expert senior technical document reviewer for automotive/embedded systems engineering.
You are reviewing a Hardware Design Document (HDD), WCCA report, or SCTM for a commercial product.

Your job is to find EVERY real issue. BE EXHAUSTIVE: report every individual issue
you find in this section — do NOT summarize, group, cap, or stop early. If a section
has 20 issues, return 20 findings. A separate verification step handles precision,
so err on the side of reporting MORE rather than fewer.

## MANDATORY CHECKS (check ALL of these):

### Text Quality:
- Spelling errors, grammar mistakes, awkward sentences
- Terminology used inconsistently (e.g., "Max Ratings" vs "Maximum Ratings" vs "Absolute Maximum Ratings")
- Abbreviations/acronyms not defined on first use
- Date format inconsistency (e.g., "13-Jun-2025" vs "10-Jul-25")
- Placeholder text, dummy text, nonsensical strings

### Technical Accuracy:
- Values without proper units
- Calculation errors or suspicious values
- Missing subscript/superscript in variable names (COUT instead of C_OUT)
- Component references (resistors, capacitors, ICs) that seem inconsistent
- Orphan references copied from datasheets (e.g., "(Note 1)", "See Figure 9-4") that don't exist in this document
- Missing explanations for suddenly-appearing values or calculations

### Cross-References & Structure:
- Figure/Table/Equation references that point to non-existent items
- Section headings that are inconsistent in format
- Incomplete sentences or thoughts

## Review Categories to use:
{cat_list}
{glossary_block}
## Document Context (structural summary):
{doc_summary[:6000]}

## Section to Review (Chunk {chunk_num}):
{chunk_text}

## Output Rules:
1. Use the [Page X] markers from the text to determine the page number
2. The `evidence` field is MANDATORY — copy-paste the exact verbatim text from the document that contains the error (5–80 chars). Do NOT leave it empty.
3. Be specific — don't say "there might be an issue", say exactly what's wrong
4. `fix_type`: "AUTO" for simple text replacements (spelling), "MANUAL" for everything else

## Output Format:
Return ONLY a JSON array. Each finding:
{{
  "category": "CATEGORY_ID",
  "severity": "CRITICAL|MAJOR|MINOR",
  "page": "page number from markers",
  "section": "section heading or reference",
  "comment": "description of the error — what is wrong and why",
  "fix": "Step-by-step fix instruction",
  "fix_type": "AUTO|MANUAL",
  "evidence": "EXACT verbatim text from the document proving the issue (required)"
}}

If no issues found, return [].
"""

    # Content-addressed cache: identical (model, chunk, categories, glossary) →
    # identical findings with no LLM call. Opt-in via DOCAI_CACHE_DIR.
    cache_key = _cache_key("chunk", model, chunk_text, ",".join(active_categories), glossary_text, seed)
    cached = _cache_get(cache_key)
    if cached is not None:
        return cached

    # Deterministic, schema-constrained call. num_predict raised so a dense chunk
    # with many issues isn't truncated mid-JSON (which silently dropped findings).
    findings = _chat_findings(client, model, prompt, f"llm_chunk_{chunk_num}", num_predict=8192, seed=seed)
    _cache_put(cache_key, findings)
    return findings


def _review_consistency_with_llm(client, model, doc_summary, glossary_text=""):
    """
    Check cross-document consistency issues.

    For large documents the summary can exceed 200K chars; we send a structured
    condensed form (section headings + statistics + first meaningful body content)
    rather than a raw character slice so the model always sees the full structure
    regardless of document size.
    """
    glossary_block = f"\n{glossary_text}\n" if glossary_text else ""

    # Build a dense structural summary: header lines (stats, headings) first,
    # then as much body content as fits — ensures coverage of the whole document.
    lines = doc_summary.split("\n")
    header_lines = [l for l in lines if l.startswith("#") or l.startswith("Total") or l.startswith("Default")]
    body_lines = [l for l in lines if l not in header_lines and l.strip()]
    # Pack: all headers + body content up to 24000 chars total
    structured = "\n".join(header_lines) + "\n\n" + "\n".join(body_lines)
    structured = structured[:24000]

    prompt = f"""You are a senior technical document reviewer. Analyze the overall structure and consistency of this document.
{glossary_block}
## Full Document Summary:
{structured}

## You MUST check for ALL of the following:

1. **Cross-reference accuracy**: Do figure/table/section references point to correct items?
2. **Terminology consistency**: Is the same thing called different names in different sections? (e.g., "VDD" vs "VCC" for the same rail, "Max Ratings" vs "Maximum Ratings")
3. **Expression consistency**: Are numbers, pin lists, ranges expressed the same way? (e.g., "pins 8-18" vs "pins 8, 9, 10, ...")
4. **Date format consistency**: Are dates in the same format throughout? (e.g., "13-Jun-2025" vs "10-Jul-25")
5. **Chapter naming consistency**: Do all chapters follow the same naming pattern? (e.g., some include part numbers but others don't)
6. **Abbreviation definitions**: Are abbreviations defined on first use?
7. **Key component coverage**: Are components mentioned in text also listed in component tables/BOM?
8. **Logical flow**: Does each technical section have a clear premise → calculation → conclusion?
9. **Placeholder/dummy text**: Any nonsensical strings or incomplete content?

## Output Format:
Return a JSON array of findings. Each finding must be:
{{
  "category": "CATEGORY_ID",
  "severity": "CRITICAL|MAJOR|MINOR",
  "page": "page number or 'ALL'",
  "section": "section reference or 'ALL'",
  "comment": "description of the inconsistency with specific examples",
  "fix": "how to resolve the issue",
  "fix_type": "MANUAL"
}}

Use these categories: TERMINOLOGY_CONSISTENCY, CROSS_REFERENCE_ACCURACY, LOGICAL_CONSISTENCY, FORMATTING_ALIGNMENT, SIGNAL_VARIABLE_NAMING, DATASHEET_COPY_ERROR

Return ONLY the JSON array. If no issues, return [].
"""

    return _chat_findings(client, model, prompt, "llm_consistency")


def review_cross_document(client, model, parsed_docs, glossary_text=""):
    """
    Cross-document review for a SUITE (e.g. SDD + SCTM + test report).

    parsed_docs: list of parsed-document dicts (each from parse_document/parse_excel).
    Finds issues that are invisible when reviewing one file at a time:
      - the same signal/pin/part/rail named differently across documents
      - values that disagree between documents
      - requirements in one document not covered/traced in another
    """
    if not parsed_docs or len(parsed_docs) < 2:
        return []

    from doc_parser import get_document_summary
    blocks = []
    for d in parsed_docs:
        name = d.get("filename", "document")
        summ = get_document_summary(d)
        # Keep each doc's contribution bounded so the suite fits one prompt.
        blocks.append(f"===== DOCUMENT: {name} =====\n{summ[:9000]}")
    suite_text = "\n\n".join(blocks)[:60000]

    glossary_block = f"\n{glossary_text}\n" if glossary_text else ""
    prompt = f"""You are reviewing a SUITE of related engineering documents together (e.g. design doc, SCTM, test report). Find issues that only appear ACROSS documents.
{glossary_block}
Find ONLY cross-document issues:
- The same signal / pin / net / part number / power rail named differently in different documents (e.g. 'HVDCDC' in one, 'HV DCDC' in another)
- A value or rating that disagrees between documents (e.g. SDD says 3.3 V, test report says 5 V for the same net)
- A requirement stated in one document but not covered/traced in the SCTM or test report
- A component listed in one document but missing from another where it should appear

Do NOT report issues internal to a single document — only cross-document mismatches.

## Documents in the suite:
{suite_text}

Return ONLY a JSON array. Each finding MUST name BOTH documents involved in `section` and quote the conflicting text in `evidence`.
[{{"category":"TERMINOLOGY_CONSISTENCY|CROSS_REFERENCE_ACCURACY|TRACEABILITY|LOGICAL_CONSISTENCY","severity":"CRITICAL|MAJOR|MINOR","page":"-","section":"DocA vs DocB","comment":"the cross-document mismatch","fix":"how to reconcile","fix_type":"MANUAL","evidence":"conflicting text"}}]
If none, return [].
"""
    return _chat_findings(client, model, prompt, "llm_cross_document", num_predict=4096)


# ============================================================
# CATEGORY-SPECIALIZED PASSES (B2) — focused prompts, higher recall per domain
# ============================================================
_UNIT_TOKENS = ("V", "mV", "kV", "A", "mA", "uA", "µA", "W", "mW", "kW",
                "Hz", "kHz", "MHz", "GHz", "Ω", "kΩ", "MΩ", "°C", "°",
                "uF", "µF", "nF", "pF", "nH", "uH", "µH", "mH", "%", "ppm", "dB", "ns", "us", "µs", "ms")


def _review_units_pass(client, model, parsed_doc, glossary_text=""):
    """Focused pass — ONLY units, calculations, and dimensional consistency."""
    parts = []
    for tbl in parsed_doc.get("tables", []):
        name = tbl.get("name", f"Table {tbl.get('index', 0) + 1}")
        rows = "\n".join("  " + " | ".join(c[:120] for c in row) for row in tbl.get("rows", []))
        parts.append(f"[{name}]\n{rows}")
    for section in parsed_doc.get("sections", []):
        for para in section.get("paragraphs", []):
            t = para.get("text", "")
            if re.search(r"\d", t) and any(u in t for u in _UNIT_TOKENS):
                parts.append(t)
    context = "\n".join(parts)
    if not context.strip():
        return []
    glossary_block = f"\n{glossary_text}\n" if glossary_text else ""
    prompt = f"""You are an electronics/power-engineering reviewer. Focus EXCLUSIVELY on UNITS, CALCULATIONS, and DIMENSIONAL CONSISTENCY. Ignore grammar, formatting, and everything else.
{glossary_block}
Find ONLY:
- Dimensionally invalid statements (e.g. "T = 180 kHz = 12.5 us" — a frequency equated to a time)
- Calculation errors / results that don't follow from the inputs
- Values missing a unit where one is required
- Inconsistent units for the same quantity (e.g. mixing mA and A in one series)
- Impossible / out-of-range values (negative resistance, efficiency > 100%, etc.)

## Numeric content from the document:
{context[:30000]}

Return ONLY a JSON array. Every finding MUST include an exact `evidence` quote.
[{{"category":"UNITS_CALCULATIONS","severity":"CRITICAL|MAJOR|MINOR","page":"-","section":"...","comment":"what is wrong and why","fix":"...","fix_type":"MANUAL","evidence":"exact text"}}]
If none, return [].
"""
    return _chat_findings(client, model, prompt, "llm_units")


def _review_terminology_pass(client, model, parsed_doc, glossary_text=""):
    """Focused pass — ONLY terminology consistency and signal/pin naming drift."""
    raw_text = parsed_doc.get("raw_text", "")
    if not raw_text.strip():
        return []
    glossary_block = f"\n{glossary_text}\n" if glossary_text else ""
    prompt = f"""You are a technical-document reviewer. Focus EXCLUSIVELY on TERMINOLOGY and SIGNAL/PIN NAMING. Ignore grammar, units, and formatting.
{glossary_block}
Find ONLY:
- The same thing called different names (e.g. "HVDCDC" vs "HV DCDC" vs "HVDC"; "VDD" vs "VCC" for one rail)
- Signal/pin/component names spelled differently across the document (e.g. "OUTAL" vs "OUT_AL")
- Acronyms used before they are defined
- Part numbers referenced inconsistently

## Document text:
{raw_text[:40000]}

Return ONLY a JSON array. Every finding MUST include an exact `evidence` quote showing BOTH variants where possible.
[{{"category":"TERMINOLOGY_CONSISTENCY","severity":"MAJOR|MINOR","page":"-","section":"...","comment":"which terms conflict","fix":"...","fix_type":"MANUAL","evidence":"exact text"}}]
If none, return [].
"""
    return _chat_findings(client, model, prompt, "llm_terminology")


def _review_tables_with_llm(client, model, parsed_doc, active_categories=None):
    """Enhanced table review — checks decimal consistency, units, completeness."""
    if not parsed_doc["tables"]:
        return []
    
    if not active_categories:
        active_categories = ["UNITS_CALCULATIONS", "TEST_RESULT_COMPLETENESS", "MEASUREMENT_RESOLUTION", 
                            "FORMATTING_ALIGNMENT", "DECIMAL_CONSISTENCY", "TABLE_QUALITY"]

    # Process ALL tables — every single table gets reviewed, no exceptions.
    # Large documents (100+ tables) are split into batches so the prompt stays
    # within context limits, but every batch is sent independently.
    all_tables = parsed_doc["tables"]
    BATCH_SIZE = 10  # tables per LLM call — each call covers all rows in its batch
    batches = [all_tables[i:i+BATCH_SIZE] for i in range(0, len(all_tables), BATCH_SIZE)]
    all_table_findings = []

    def _build_table_prompt(batch, batch_idx):
        tables_text = []
        for tbl in batch:
            tbl_name = tbl.get("name", f"Table {tbl['index'] + 1}")
            rows_str = "\n".join(
                f"  Row {i}: {' | '.join(c[:400] for c in row)}"
                for i, row in enumerate(tbl["rows"])  # ALL rows, no cap
            )
            tables_text.append(f"--- {tbl_name} ({tbl['num_rows']}×{tbl['num_cols']}) ---\n{rows_str}")

        return f"""You are reviewing TABLES (batch {batch_idx}/{len(batches)}) in an engineering technical document (HDD/WCCA/SCTM). Check EVERY table for:

1. **Decimal place consistency**: Within each column, do all numerical values have the same number of decimal places?
2. **Missing headers or unclear column names**
3. **Empty cells that should have values**
4. **Inconsistent units across rows** (e.g., some cells say "V" and others say "VDC")
5. **Test results without pass/fail criteria, actual measurements, or judgments**
6. **Numerical values without units**
7. **Suspicious values** (e.g., a voltage column has a value of "100" among "3.3", "5.0", "12.0")
8. **Missing legends** — if a table uses symbols or abbreviations, is there a legend?
9. **Min/Typ/Max values in wrong columns** compared to datasheet
10. **Placeholder text or dummy data**

IMPORTANT: Do NOT report that a table is "incomplete" or "truncated" — data may continue on another page.

## Tables:
{chr(10).join(tables_text)}

## Output Format:
Return a JSON array. The `evidence` field is MANDATORY — copy the exact cell value or header that shows the issue.
{{
  "category": "CATEGORY_ID",
  "severity": "CRITICAL|MAJOR|MINOR",
  "page": "-",
  "section": "Exact Table Name",
  "comment": "detailed description with specific cell/row references",
  "fix": "step-by-step fix instruction",
  "fix_type": "MANUAL",
  "evidence": "exact quoted text from the table that shows the problem"
}}

Return ONLY the JSON array. If no issues, return [].
"""

    def _run_batch(args):
        batch_idx, batch = args
        prompt = _build_table_prompt(batch, batch_idx)
        try:
            return _chat_findings(client, model, prompt, f"llm_tables_batch{batch_idx}", num_predict=8192)
        except Exception as e:
            print(f"[tables] batch {batch_idx}/{len(batches)} failed: {e}")
            return []

    # Run table batches concurrently (same knob as the chunk pass).
    import concurrent.futures as _cf
    concurrency = max(1, int(os.environ.get("DOCAI_CONCURRENCY", "5")))
    indexed = list(enumerate(batches, 1))
    if concurrency == 1 or len(indexed) <= 1:
        for args in indexed:
            all_table_findings.extend(_run_batch(args))
    else:
        with _cf.ThreadPoolExecutor(max_workers=concurrency) as pool:
            for res in pool.map(_run_batch, indexed):
                all_table_findings.extend(res)

    return all_table_findings


def _review_images_with_llm(client, model, parsed_doc, doc_summary="", progress_callback=None, errors=None):
    """Review images/diagrams using a Vision model, cross-referencing with document text."""
    if not parsed_doc.get("images"):
        return []

    findings = []
    
    # Build a mapping of nearby text for each image for cross-referencing
    all_text_by_section = {}
    for section in parsed_doc.get("sections", []):
        heading = section.get("heading", "Unknown")
        paragraphs_text = " ".join(p["text"] for p in section.get("paragraphs", []) if p.get("text"))
        if paragraphs_text:
            all_text_by_section[heading] = paragraphs_text  # full section text for image cross-ref
    
    # Build surrounding text context — all sections, full text (no per-section slice)
    nearby_text = "\n".join([f"[{h}]: {t}" for h, t in all_text_by_section.items()])

    # Review ALL non-trivial images — no arbitrary cap.
    # Per-image errors are caught individually so one failure never stops the rest.
    reviewable_images = [img for img in parsed_doc["images"] if img.get("full_b64") and not img.get("is_small")]
    total_images = len(reviewable_images)

    def _review_one_image(idx, img):
        prompt = f"""You are reviewing a DIAGRAM / GRAPH / IMAGE in a technical engineering document.

You have TWO jobs:

**JOB 1 — Image Quality Check:**
1. Missing labels, axis titles, or unclear legends
2. Unreadable text within the diagram
3. Misspelled words or typos visible in the image
4. Formatting issues, cropped edges, or low resolution
5. Dummy/placeholder images or nonsense text

**JOB 2 — Cross-Check Image vs Document Text:**
The document text near this image is provided below. Compare the image content against the text:
6. Does the image match what the text describes? (e.g., if text says "output voltage is 3.3V" but the graph shows 5V, that's a CRITICAL error)
7. Are values/measurements in the image consistent with the document text?
8. Are component names, signal names, or labels in the image consistent with the text?
9. Is the image referenced properly in the text?

## Surrounding Document Text (for cross-reference):
{nearby_text[:3000]}

## Document Summary:
{doc_summary[:4000]}

## Output Format:
Return a JSON array of findings. Each finding must be:
{{
  "category": "WAVEFORM_DOCUMENTATION",
  "severity": "CRITICAL|MAJOR|MINOR",
  "page": "-",
  "section": "Image / Diagram",
  "comment": "detailed description of issue found",
  "fix": "step-by-step instruction on how to fix this",
  "fix_type": "MANUAL"
}}

IMPORTANT: Use severity CRITICAL if image data contradicts the document text. Use MAJOR for missing labels. Use MINOR for formatting issues.
Return ONLY the JSON array. If no issues, return [].
"""
        # Per-image try/except keeps one bad image from aborting the whole pass;
        # failures are recorded so the orchestrator can surface them.
        try:
            img_findings = _chat_findings(
                client, model, prompt, f"llm_image_{idx}",
                images=[img["full_b64"]], num_predict=2048,
            )
            for f in img_findings:
                if f.get("section") in ("-", "Image / Diagram"):
                    f["section"] = f"Image {idx + 1}"
                f["image_ref"] = idx + 1
            return img_findings
        except Exception as e:
            if errors is not None:
                errors.append(f"Image {idx + 1}: {str(e)[:160]}")
            print(f"Error during image {idx} review: {str(e)}")
            return []

    # Vision calls are independent → run them concurrently (same DOCAI_CONCURRENCY
    # knob as the text/table passes). 177 images one-at-a-time was the slowest pass.
    import concurrent.futures as _cf
    concurrency = max(1, int(os.environ.get("DOCAI_CONCURRENCY", "5")))
    done = 0
    if concurrency == 1 or total_images <= 1:
        for idx, img in enumerate(reviewable_images):
            findings.extend(_review_one_image(idx, img))
            done += 1
            if progress_callback:
                progress_callback(f"Analyzing image {done}/{total_images} with vision AI...", 82 + int((done / max(total_images, 1)) * 10))
    else:
        with _cf.ThreadPoolExecutor(max_workers=concurrency) as pool:
            futs = {pool.submit(_review_one_image, idx, img): idx for idx, img in enumerate(reviewable_images)}
            for fut in _cf.as_completed(futs):
                findings.extend(fut.result())
                done += 1
                if progress_callback:
                    progress_callback(f"Analyzing images {done}/{total_images} with vision AI...", 82 + int((done / max(total_images, 1)) * 10))

    return findings


# ============================================================
# FINDING PARSERS
# ============================================================
def _parse_llm_findings(llm_response, source="llm"):
    """Parse LLM response into structured findings."""
    # Log raw output for debugging
    try:
        debug_log_path = os.path.join(os.path.dirname(__file__), "uploads", "raw_llm_responses_debug.txt")
        with open(debug_log_path, "a", encoding="utf-8") as f:
            f.write(f"\n--- NEW RESPONSE ({source}) ---\n{llm_response}\n")
    except Exception:
        pass

    text = llm_response.strip()

    # Remove markdown code blocks if present
    if "```json" in text:
        text = text.split("```json")[-1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[-1].split("```")[0].strip()

    valid_findings = []

    # First attempt: Try to parse as a direct JSON array
    try:
        start = text.find("[")
        end = text.rfind("]")
        if start != -1 and end != -1 and start < end:
            json_text = text[start : end + 1]
            findings = json.loads(json_text)
            if isinstance(findings, list):
                for f in findings:
                    if not isinstance(f, dict):
                        continue
                    cat = f.get("category", "GRAMMAR_SPELLING")
                    if cat not in REVIEW_CATEGORIES:
                        cat = "GRAMMAR_SPELLING"
                    sev = f.get("severity", "MINOR").upper()
                    if sev == "SUGGESTION": continue
                    if sev not in SEVERITY_LEVELS:
                        sev = "MINOR"
                    
                    valid_findings.append({
                        "category": cat,
                        "severity": sev,
                        "page": str(f.get("page", "-")),
                        "section": str(f.get("section", "-")),
                        "comment": str(f.get("comment", "")),
                        "fix": str(f.get("fix", "Review content for accuracy.")),
                        "fix_type": str(f.get("fix_type", "MANUAL")),
                        "evidence": str(f.get("evidence", "")).strip(),
                        "para_index": None,
                        "source": source,
                    })
                return valid_findings
    except Exception as e:
        print(f"Standard JSON parsing failed: {e}")

    # Second attempt: Robust Regex extraction for JSON objects (fallback)
    print("Attempting regex fallback extraction...")
    try:
        # Match anything between { and } that does not contain another {
        dict_strings = re.findall(r'\{[^{}]*\}', text, re.DOTALL)
        for ds in dict_strings:
            try:
                # Clean up trailing commas before } which LLMs sometimes output
                cleaned_ds = re.sub(r',\s*\}', '}', ds)
                f = json.loads(cleaned_ds)
                if isinstance(f, dict) and "category" in f and "comment" in f:
                    cat = f.get("category", "GRAMMAR_SPELLING")
                    if cat not in REVIEW_CATEGORIES:
                        cat = "GRAMMAR_SPELLING"
                    sev = f.get("severity", "MINOR").upper()
                    if sev == "SUGGESTION": continue
                    if sev not in SEVERITY_LEVELS:
                        sev = "MINOR"

                    valid_findings.append({
                        "category": cat,
                        "severity": sev,
                        "page": str(f.get("page", "-")),
                        "section": str(f.get("section", "-")),
                        "comment": str(f.get("comment", "")),
                        "fix": str(f.get("fix", "Review content for accuracy.")),
                        "fix_type": str(f.get("fix_type", "MANUAL")),
                        "evidence": str(f.get("evidence", "")).strip(),
                        "para_index": None,
                        "source": source + "_regex",
                    })
            except Exception:
                continue
    except Exception as e:
        print(f"Regex parsing failed: {e}")

    return valid_findings


# ============================================================
# EVIDENCE GROUNDING & ANCHORING (ported/extended from kimi analyzer)
# ============================================================
# Phrases that signal the model decided something is FINE — never a real finding.
NON_ERROR_PHRASES = [
    "actually correct", "is correct", "spelled correctly", "not found",
    "acceptable", "adequate", "consistent between", "no issue", "not an issue",
    "no error", "appears correct", "looks correct",
]


def _norm_text(s):
    """Lowercase, collapse whitespace — for substring/token matching."""
    if not s:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def _build_doc_index(parsed_doc):
    """
    Build (full normalized doc text, paragraph anchor list) once.
    Anchors let us attach para_index/page to a grounded finding so the
    side-by-side viewer can jump straight to the offending text.
    """
    parts = [parsed_doc.get("raw_text", "")]
    for tbl in parsed_doc.get("tables", []):
        parts.append(tbl.get("name", ""))
        for row in tbl.get("rows", []):
            parts.append(" ".join(row))
    doc_text = _norm_text("\n".join(p for p in parts if p))

    anchors = []  # {para_index, page, section, norm}
    for section in parsed_doc.get("sections", []):
        heading = section.get("heading", "")
        for para in section.get("paragraphs", []):
            text = para.get("text", "")
            if not text:
                continue
            anchors.append({
                "para_index": para.get("index"),
                "page": para.get("page"),
                "section": heading,
                "norm": _norm_text(text),
            })
    return doc_text, anchors


def _ground_and_anchor_findings(findings, parsed_doc):
    """
    Reject hallucinated LLM findings and anchor the survivors to a paragraph.

    Rules (applied only to llm-sourced findings; local findings are trusted):
      - Must carry a non-empty `evidence` quote.
      - The evidence must actually appear in the document (token overlap),
        otherwise it is a hallucination and is dropped.
      - Findings whose text says the content is fine are dropped.
    Anchoring: locate the evidence inside a paragraph and copy that
    paragraph's para_index/page onto the finding when missing.
    """
    doc_text, anchors = _build_doc_index(parsed_doc)
    kept = []

    for f in findings:
        f.setdefault("evidence", "")
        f.setdefault("para_index", None)
        source = f.get("source", "")
        is_llm = source.startswith("llm") or source.startswith("vision")

        blob = _norm_text(" ".join([f.get("comment", ""), f.get("fix", ""), f.get("evidence", "")]))
        if any(phrase in blob for phrase in NON_ERROR_PHRASES):
            continue

        if is_llm:
            f.setdefault("grounded", True)
            evidence = f.get("evidence", "").strip()
            # Image/vision findings describe a picture, not quotable text — exempt.
            is_image = source.startswith("vision") or "image" in source or f.get("image_ref")
            if not is_image:
                # When evidence is empty, try to salvage a quote from the comment.
                if not evidence:
                    comment = f.get("comment", "")
                    m = re.search(r'"([^"]{8,120})"|\'([^\']{8,120})\'|"([^"]{8,120})"', comment)
                    if m:
                        evidence = (m.group(1) or m.group(2) or m.group(3)).strip()
                        f["evidence"] = evidence
                    else:
                        evidence = comment[:60].strip()

                # RECALL-FIRST: we no longer DROP ungrounded findings — we keep them
                # and mark grounded=False so the confidence score reflects the lower
                # certainty. The user filters with the confidence slider; Max mode
                # applies a hard precision cut. This is why a 100-issue doc no longer
                # collapses to ~15 findings.
                ev_norm = _norm_text(evidence)
                tokens = [t for t in re.findall(r"[a-z0-9]{4,}", ev_norm) if len(t) > 3]
                if not evidence or (tokens and not any(tok in doc_text for tok in tokens[:4])):
                    f["grounded"] = False

                # Anchor to the first paragraph that contains the evidence snippet.
                snippet = ev_norm[:40]
                if snippet:
                    for a in anchors:
                        if snippet in a["norm"]:
                            if f.get("para_index") is None:
                                f["para_index"] = a["para_index"]
                            if str(f.get("page", "-")) in ("-", "", "ALL", "None"):
                                f["page"] = str(a["page"]) if a["page"] is not None else f.get("page", "-")
                            break

        kept.append(f)

    return kept


def _critic_filter_findings(client, model, findings, progress_callback=None, batch_size=15, errors=None):
    """
    Second-opinion "critic" pass that cuts false positives.

    Each LLM-sourced candidate is shown back to the model with its evidence and
    judged keep/drop + confidence. Local (deterministic) findings are trusted and
    assigned confidence 1.0. If the critic call fails we keep the batch (never lose
    findings to an infra error) at a neutral confidence.
    """
    llm_findings = [f for f in findings if f.get("source", "").startswith(("llm", "vision"))]
    other_findings = [f for f in findings if not f.get("source", "").startswith(("llm", "vision"))]
    for f in other_findings:
        f.setdefault("confidence", 1.0)

    if not llm_findings:
        return findings

    if progress_callback:
        progress_callback("Verifying findings (critic pass)...", None)

    verified = []
    for start in range(0, len(llm_findings), batch_size):
        batch = llm_findings[start:start + batch_size]
        items = [
            {
                "id": i,
                "category": f.get("category", ""),
                "claim": f.get("comment", "")[:800],
                "evidence": f.get("evidence", "")[:600],
            }
            for i, f in enumerate(batch)
        ]
        prompt = f"""You are a STRICT QA reviewer verifying candidate findings from an
automated technical-document review. For each candidate decide whether it is a
GENUINE defect that a senior reviewer would keep.

Set "keep": false when the candidate is vague, speculative, not actually an error,
or when the evidence does not clearly support the claim. Set "keep": true only for
real, defensible defects. Always include a confidence between 0 and 1.

Candidates (JSON):
{json.dumps(items, ensure_ascii=False)}

Return ONLY a JSON array of verdicts: {{"id": <int>, "keep": <bool>, "confidence": <0-1>, "reason": "<short>"}}.
"""
        try:
            response = _chat_with_retry(
                client.chat,
                model=model,
                messages=[{"role": "user", "content": prompt}],
                format=VERDICTS_JSON_SCHEMA,
                options=_llm_options(2048),
            )
            reply = response["message"]["content"] if isinstance(response, dict) else response.message.content
            verdicts = _parse_json_array(reply)
            verdict_by_id = {v["id"]: v for v in verdicts if isinstance(v, dict) and "id" in v}
            for i, f in enumerate(batch):
                v = verdict_by_id.get(i)
                if v is None:
                    f.setdefault("confidence", 0.5)
                elif v.get("keep", True):
                    f["confidence"] = round(float(v.get("confidence", 0.7)), 2)
                else:
                    # RECALL-FIRST: the critic no longer DELETES findings. A "drop"
                    # verdict just lowers confidence so the user can hide it with the
                    # slider (or Max mode cuts it). Nothing is silently lost.
                    f["confidence"] = min(round(float(v.get("confidence", 0.25)), 2), 0.35)
                verified.append(f)
        except Exception as e:
            if errors is not None:
                errors.append(f"critic: {str(e)[:140]}")
            for f in batch:
                f.setdefault("confidence", 0.6)
            verified.extend(batch)

    return other_findings + verified


def _parse_json_array(text):
    """Best-effort parse of a JSON array from an LLM reply (schema-constrained)."""
    if not text:
        return []
    text = text.strip()
    if "```" in text:
        text = text.split("```")[-2] if text.count("```") >= 2 else text.replace("```", "")
        text = text.replace("json", "", 1).strip()
    start, end = text.find("["), text.rfind("]")
    if start == -1 or end == -1 or end <= start:
        return []
    try:
        data = json.loads(text[start:end + 1])
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _deduplicate_findings(findings):
    """
    Remove duplicate / near-duplicate findings, including the same issue reported
    by different sources (e.g. decimal consistency found by both a local check and
    the table LLM pass). Two same-category findings are duplicates when:
      - their comments are >70% similar, OR
      - they share a location (section + page) and comments are >50% similar, OR
      - one's evidence is contained in the other's and they share a location.
    """
    unique = []
    for f in findings:
        f_comment = _norm_text(f.get("comment", ""))
        f_section = _norm_text(f.get("section", ""))
        f_page = str(f.get("page", "-"))
        f_ev = _norm_text(f.get("evidence", ""))
        is_duplicate = False
        for u in unique:
            if f.get("category") != u.get("category"):
                continue
            u_comment = _norm_text(u.get("comment", ""))
            same_loc = (f_section == _norm_text(u.get("section", ""))) and (f_page == str(u.get("page", "-")))
            ratio = difflib.SequenceMatcher(None, f_comment, u_comment).ratio()
            u_ev = _norm_text(u.get("evidence", ""))
            ev_overlap = bool(f_ev and u_ev and (f_ev in u_ev or u_ev in f_ev))
            if ratio > 0.7 or (same_loc and ratio > 0.5) or (same_loc and ev_overlap):
                is_duplicate = True
                # Prefer to keep the higher-severity / higher-confidence copy.
                if SEVERITY_LEVELS.get(f.get("severity", "MINOR"), {}).get("weight", 0) > \
                   SEVERITY_LEVELS.get(u.get("severity", "MINOR"), {}).get("weight", 0):
                    u.update(f)
                break
        if not is_duplicate:
            unique.append(f)
    return unique


def _findings_match(a, b):
    """True if two findings describe the same issue (same category + similar text/evidence)."""
    if a.get("category") != b.get("category"):
        return False
    a_ev, b_ev = _norm_text(a.get("evidence", "")), _norm_text(b.get("evidence", ""))
    if a_ev and b_ev and (a_ev in b_ev or b_ev in a_ev):
        return True
    ratio = difflib.SequenceMatcher(None, _norm_text(a.get("comment", "")), _norm_text(b.get("comment", ""))).ratio()
    return ratio > 0.6


def diff_findings(old_findings, new_findings):
    """
    Compare two reviews of (different versions of) a document.

    Returns {"new": [...], "fixed": [...], "unchanged": [...]}:
      - new       — in the new review but not the old (newly introduced issues)
      - fixed     — in the old review but not the new (resolved since last time)
      - unchanged — present in both
    """
    old = list(old_findings or [])
    new = list(new_findings or [])
    matched_old = set()
    result = {"new": [], "fixed": [], "unchanged": []}

    for nf in new:
        hit = None
        for i, of in enumerate(old):
            if i in matched_old:
                continue
            if _findings_match(nf, of):
                hit = i
                break
        if hit is not None:
            matched_old.add(hit)
            result["unchanged"].append(nf)
        else:
            result["new"].append(nf)

    for i, of in enumerate(old):
        if i not in matched_old:
            result["fixed"].append(of)

    return result


# ============================================================
# TOC & HEADING STRUCTURE CHECKS
# ============================================================

def _check_toc_heading_sync(parsed_doc):
    """
    Compare TOC entries against actual document headings.
    Flags: stale TOC text, numbering mismatches, orphan TOC rows,
    and headings missing from the TOC.

    Fix (I1): uses max_toc_level so headings intentionally omitted from
    the TOC (e.g. level-3 when TOC only goes to level-2) are not
    falsely flagged as missing.

    Fix (I2): uses id(candidate) instead of candidate["index"] so two
    heading objects that happen to share a paragraph index are tracked
    independently.
    """
    findings = []
    toc_entries = [e for e in parsed_doc.get("toc", {}).get("entries", []) if e.get("text")]
    headings = [h for h in parsed_doc.get("headings", []) if h.get("text")]

    if not toc_entries or not headings:
        return findings

    # Only compare headings up to the deepest level present in the TOC.
    # This prevents false "missing from TOC" findings for intentionally
    # omitted deeper levels (e.g. Heading 3 when TOC only shows 2 levels).
    max_toc_level = max((e.get("level") or 1 for e in toc_entries), default=1)
    comparable_headings = [h for h in headings if (h.get("level") or 0) <= max_toc_level]
    if not comparable_headings:
        return findings

    actual_by_number = defaultdict(list)
    actual_by_title = defaultdict(list)
    for heading in comparable_headings:
        title_key = _normalize_toc_text(heading.get("title") or heading.get("text"))
        if heading.get("number"):
            actual_by_number[heading["number"]].append(heading)
        actual_by_title[title_key].append(heading)

    # Track which headings were matched using object identity (avoids index collisions)
    matched_heading_ids = set()

    for entry in toc_entries:
        entry_text = entry.get("text", "")
        entry_title = entry.get("title") or entry_text
        entry_title_key = _normalize_toc_text(entry_title)
        entry_number = entry.get("number")
        entry_level = entry.get("level") or 1
        candidate = None

        if entry_number:
            same_number = actual_by_number.get(entry_number, [])
            exact_match = [
                h for h in same_number
                if _normalize_toc_text(h.get("title") or h.get("text")) == entry_title_key
            ]
            if exact_match:
                candidate = exact_match[0]
            elif same_number:
                candidate = same_number[0]
                findings.append({
                    "category": "TOC_VALIDATION",
                    "severity": "MAJOR",
                    "page": str(entry.get("page", candidate.get("page", "-"))),
                    "section": "Table of Contents",
                    "comment": (
                        f"TOC entry '{entry_text}' does not match the actual heading text "
                        f"for section {entry_number}. The document heading reads '{candidate.get('text')}'."
                    ),
                    "fix": f"Update the TOC entry for section {entry_number} to match '{candidate.get('text')}'.",
                    "source": "local",
                    "fix_type": "MANUAL",
                })
            else:
                same_title = actual_by_title.get(entry_title_key, [])
                if same_title:
                    candidate = same_title[0]
                    findings.append({
                        "category": "TOC_VALIDATION",
                        "severity": "MAJOR",
                        "page": str(entry.get("page", candidate.get("page", "-"))),
                        "section": "Table of Contents",
                        "comment": (
                            f"TOC entry '{entry_text}' uses section number {entry_number}, "
                            f"but the matching heading is numbered "
                            f"'{candidate.get('number') or 'unnumbered'}'."
                        ),
                        "fix": f"Correct the TOC numbering for '{candidate.get('text')}'.",
                        "source": "local",
                        "fix_type": "MANUAL",
                    })
                else:
                    findings.append({
                        "category": "TOC_VALIDATION",
                        "severity": "MAJOR",
                        "page": str(entry.get("page", "-")),
                        "section": "Table of Contents",
                        "comment": (
                            f"TOC entry '{entry_text}' does not map to any heading in the "
                            f"document body. The TOC may be stale or contain an extra entry."
                        ),
                        "fix": f"Remove or update the TOC entry '{entry_text}' to match a real heading.",
                        "source": "local",
                        "fix_type": "MANUAL",
                    })
        else:
            same_title = actual_by_title.get(entry_title_key, [])
            if same_title:
                candidate = same_title[0]
            else:
                findings.append({
                    "category": "TOC_VALIDATION",
                    "severity": "MAJOR",
                    "page": str(entry.get("page", "-")),
                    "section": "Table of Contents",
                    "comment": f"TOC entry '{entry_text}' does not match any heading in the document body.",
                    "fix": f"Update or remove the TOC entry '{entry_text}'.",
                    "source": "local",
                    "fix_type": "MANUAL",
                })

        if candidate:
            matched_heading_ids.add(id(candidate))
            if candidate.get("level") != entry_level:
                findings.append({
                    "category": "TOC_VALIDATION",
                    "severity": "MINOR",
                    "page": str(candidate.get("page", "-")),
                    "section": "Table of Contents",
                    "comment": (
                        f"TOC nesting mismatch for '{candidate.get('text')}'. "
                        f"TOC uses level {entry_level} but the actual heading is level {candidate.get('level')}."
                    ),
                    "fix": f"Adjust the TOC indent level for '{candidate.get('text')}' to match the heading hierarchy.",
                    "source": "local",
                    "fix_type": "MANUAL",
                })

    # Flag headings missing from the TOC
    for heading in comparable_headings:
        if id(heading) in matched_heading_ids:
            continue
        findings.append({
            "category": "TOC_VALIDATION",
            "severity": "MAJOR",
            "page": str(heading.get("page", "-")),
            "section": heading.get("text", "Unknown"),
            "comment": f"Heading '{heading.get('text')}' appears in the document body but is missing from the TOC.",
            "fix": f"Refresh the TOC so it includes '{heading.get('text')}'.",
            "source": "local",
            "fix_type": "MANUAL",
        })

    return findings


def _check_section_number_continuity(parsed_doc):
    """
    Verify numbered headings are sequential within each parent branch.
    Flags gaps (1.1, 1.3 — missing 1.2) and duplicates (two sections both 2.2).
    """
    findings = []
    numbered_headings = []

    for heading in parsed_doc.get("headings", []):
        number = heading.get("number")
        if not number:
            continue
        try:
            parts = tuple(int(p) for p in number.split("."))
        except ValueError:
            continue  # e.g. annex A.1, B.2 — skip non-integer numbering
        numbered_headings.append({"heading": heading, "parts": parts})

    siblings = defaultdict(list)
    for item in numbered_headings:
        parent = item["parts"][:-1]
        siblings[parent].append(item)

    for parent, items in siblings.items():
        child_values = [item["parts"][-1] for item in items]

        # Flag duplicates
        duplicates = [v for v, cnt in Counter(child_values).items() if cnt > 1]
        for dup_val in sorted(duplicates):
            dup_items = [item for item in items if item["parts"][-1] == dup_val]
            anchor = dup_items[1]["heading"]
            dup_label = _format_section_number(parent + (dup_val,))
            findings.append({
                "category": "TOC_VALIDATION",
                "severity": "MAJOR",
                "page": str(anchor.get("page", "-")),
                "section": anchor.get("text", "Unknown"),
                "comment": f"Duplicate section number '{dup_label}' detected in the heading sequence.",
                "fix": f"Rename one of the duplicate headings numbered '{dup_label}' so numbering is unique.",
                "source": "local",
                "fix_type": "MANUAL",
            })

        # Flag gaps in the sequence
        unique_values = sorted(set(child_values))
        for prev, curr in zip(unique_values, unique_values[1:]):
            if curr - prev <= 1:
                continue
            missing = [_format_section_number(parent + (v,)) for v in range(prev + 1, curr)]
            impacted = next(
                (item["heading"] for item in items if item["parts"][-1] == curr),
                items[-1]["heading"],
            )
            parent_label = _format_section_number(parent) or "top level"
            findings.append({
                "category": "TOC_VALIDATION",
                "severity": "MAJOR",
                "page": str(impacted.get("page", "-")),
                "section": impacted.get("text", "Unknown"),
                "comment": (
                    f"Non-continuous section numbering under {parent_label}. "
                    f"Missing section number(s): {', '.join(missing)}."
                ),
                "fix": f"Renumber headings under {parent_label} so the sequence is continuous.",
                "source": "local",
                "fix_type": "MANUAL",
            })

    return findings


def _check_heading_hierarchy(parsed_doc):
    """
    Ensure no heading appears without its parent heading level or parent section number.
    Flags orphan sub-sections such as a Heading 3 with no preceding Heading 2.

    Fix (B3): uses elif so a heading only produces ONE orphan finding even when both
    the style-level check and the numbered-parent check would fire.
    """
    findings = []
    open_headings = {}
    seen_numbered_sections = set()

    headings = sorted(parsed_doc.get("headings", []), key=lambda h: h.get("index", 0))
    for heading in headings:
        level = heading.get("level")
        number = heading.get("number")
        text = heading.get("text", "Unknown")
        if not level:
            continue

        # Trim open_headings to only ancestors of the current level
        open_headings = {lvl: h for lvl, h in open_headings.items() if lvl < level}

        missing_parent = False
        reason = None

        # Check 1: style-level orphan (Heading 3 with no Heading 2 above)
        if level > 1 and (level - 1) not in open_headings:
            missing_parent = True
            reason = f"no preceding Heading {level - 1}"

        # Check 2 (elif — only fires if check 1 didn't): numeric-parent orphan
        elif number and "." in number:
            parent_number = number.rsplit(".", 1)[0]
            if parent_number not in seen_numbered_sections:
                missing_parent = True
                reason = f"parent section '{parent_number}' does not exist earlier in the document"

        if missing_parent:
            findings.append({
                "category": "TOC_VALIDATION",
                "severity": "MAJOR",
                "page": str(heading.get("page", "-")),
                "section": text,
                "comment": f"Orphan heading '{text}' detected: {reason}.",
                "fix": f"Add the missing parent heading before '{text}' or correct its heading level/numbering.",
                "source": "local",
                "fix_type": "MANUAL",
            })

        open_headings[level] = heading
        if number:
            seen_numbered_sections.add(number)

    return findings


def _normalize_toc_text(text):
    """Normalise heading text for deterministic TOC comparisons (case-insensitive, whitespace-collapsed)."""
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip().casefold()


def _format_section_number(parts):
    """Format a section-number tuple back to a dotted label, e.g. (1, 2) → '1.2'."""
    if not parts:
        return ""
    return ".".join(str(p) for p in parts)


def _check_spacing_errors(parsed_doc):
    findings = []
    for section in parsed_doc.get("sections", []):
        for para in section.get("paragraphs", []):
            text = para.get("text", "")
            if "  " in text.strip():
                # Avoid flagging indentation spaces
                findings.append({
                    "category": "FORMATTING_ALIGNMENT",
                    "severity": "MINOR",
                    "page": str(para.get("page", "-")),
                    "section": section.get("heading", ""),
                    "comment": "Double spaces detected within the text. Use a single space after punctuation.",
                    "fix": "Remove extra spaces.",
                    "source": "local",
                    "fix_type": "MANUAL",
                })
                break # Only flag once per section max to avoid noise
    return findings

def _check_repeated_words(parsed_doc):
    findings = []
    repeated_pattern = re.compile(r'\b([A-Za-z]+)\s+\1\b', re.IGNORECASE)
    for section in parsed_doc.get("sections", []):
        for para in section.get("paragraphs", []):
            text = para.get("text", "")
            match = repeated_pattern.search(text)
            if match:
                findings.append({
                    "category": "GRAMMAR_SPELLING",
                    "severity": "MINOR",
                    "page": str(para.get("page", "-")),
                    "section": section.get("heading", ""),
                    "comment": f"Repeated word detected: '{match.group(0)}'.",
                    "fix": f"Remove the duplicate word '{match.group(1)}'.",
                    "source": "local",
                    "fix_type": "AUTO",
                })
    return findings

def _check_unmatched_brackets(parsed_doc):
    """
    Flag unbalanced parentheses, but ONLY in self-contained paragraphs that end
    with sentence punctuation. This avoids false positives when a parenthetical
    legitimately spans multiple paragraphs/list items.
    """
    findings = []
    for section in parsed_doc.get("sections", []):
        for para in section.get("paragraphs", []):
            text = para.get("text", "")
            if not text:
                continue
            stripped = text.rstrip()
            # Only judge balance on paragraphs that look like a complete sentence.
            if not stripped.endswith((".", "!", "?", ":", ";")):
                continue
            if text.count("(") != text.count(")") or text.count("[") != text.count("]"):
                findings.append({
                    "category": "GRAMMAR_SPELLING",
                    "severity": "MINOR",
                    "page": str(para.get("page", "-")),
                    "section": section.get("heading", ""),
                    "comment": f"Unmatched bracket detected in text: '{text[:80]}'.",
                    "fix": "Ensure all brackets are opened and closed correctly.",
                    "evidence": text[:120],
                    "source": "local",
                    "fix_type": "MANUAL",
                })
    return findings

def _check_empty_paragraphs(parsed_doc):
    # Too many empty paragraphs can mean bad formatting
    findings = []
    empty_streak = 0
    for section in parsed_doc.get("sections", []):
        for para in section.get("paragraphs", []):
            if not para.get("text", "").strip():
                empty_streak += 1
                if empty_streak == 3:
                    findings.append({
                        "category": "FORMATTING_ALIGNMENT",
                        "severity": "MINOR",
                        "page": str(para.get("page", "-")),
                        "section": section.get("heading", ""),
                        "comment": "Multiple consecutive empty line breaks detected.",
                        "fix": "Use paragraph spacing instead of empty hard carriage returns.",
                        "source": "local",
                        "fix_type": "MANUAL",
                    })
            else:
                empty_streak = 0
    return findings

def _check_min_typ_max_tables(parsed_doc):
    findings = []
    # Search tables for Min, Typ, Max columns and ensure logical order
    for tbl in parsed_doc.get("tables", []):
        rows = tbl.get("rows", [])
        if not rows: continue
        headers = [col.strip().lower() for col in rows[0]]
        
        min_idx = headers.index('min') if 'min' in headers else -1
        typ_idx = headers.index('typ') if 'typ' in headers else -1
        max_idx = headers.index('max') if 'max' in headers else -1
        
        if typ_idx != -1 and ((min_idx != -1 and min_idx > typ_idx) or (max_idx != -1 and typ_idx > max_idx)):
            findings.append({
                "category": "TABLE_QUALITY",
                "severity": "MINOR",
                "page": "-",
                "section": tbl.get("name", "Table"),
                "comment": "Min, Typ, Max columns are not in standard logical order.",
                "fix": "Reorder columns to Min, Typ, Max.",
                "source": "local",
                "fix_type": "MANUAL",
            })
    return findings

def _check_unit_standardization(parsed_doc):
    """
    Flag missing spaces before units (e.g. '5V' → '5 V') but skip common
    power-rail/net names (5V, 3V3, 12V…) which are identifiers, not measurements.
    One finding per section to keep volume down.
    """
    findings = []
    bad_unit_pattern = re.compile(r'\b\d+(?:\.\d+)?(V|mA|uA|A|Hz|kHz|MHz|W|mW)\b')
    for section in parsed_doc.get("sections", []):
        flagged_here = False
        for para in section.get("paragraphs", []):
            if flagged_here:
                break
            text = para.get("text", "")
            for match in bad_unit_pattern.finditer(text):
                token = match.group(0)
                if token.upper() in COMMON_RAIL_NAMES:
                    continue  # net/rail name, not a measurement
                findings.append({
                    "category": "UNITS_CALCULATIONS",
                    "severity": "MINOR",
                    "page": str(para.get("page", "-")),
                    "section": section.get("heading", ""),
                    "comment": f"Missing space before unit: '{token}'. Convention is a space between value and unit (e.g. '{token[:-len(match.group(1))]} {match.group(1)}').",
                    "fix": "Insert a space between the numeric value and the unit.",
                    "evidence": token,
                    "source": "local",
                    "fix_type": "MANUAL",
                })
                flagged_here = True
                break
    return findings

