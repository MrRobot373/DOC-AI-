# Kimi-Style Document Analyzer

Experimental analyzer for deeper engineering-document review.

Goal: produce Kimi-like reports with stronger evidence, better categories, severity sheets, and category sheets while leaving the current app untouched.

This cannot guarantee "100% perfect" results. The practical target is:

- more complete coverage
- fewer weak/non-error findings
- explicit evidence for every finding
- repeatable local checks before LLM review
- review of all tables/images in batches instead of hard limits

## Run

Local deterministic checks only:

```powershell
python kimi_style_analyzer/analyze.py ACC_Ph3_HardwareDesignDocument_WithAIChk.docx
```

With Ollama Cloud LLM review:

```powershell
$env:OLLAMA_API_KEY="your-key"
python kimi_style_analyzer/analyze.py ACC_Ph3_HardwareDesignDocument_WithAIChk.docx --llm --model gpt-oss:120b
```

With a vision model for images:

```powershell
$env:OLLAMA_API_KEY="your-key"
python kimi_style_analyzer/analyze.py ACC_Ph3_HardwareDesignDocument_WithAIChk.docx --llm --model gpt-oss:120b --vision-model qwen2.5vl:72b
```

Output is written to `kimi_style_analyzer/output/`.

## Architecture

1. Parse DOCX using the existing repo parser.
2. Run deterministic local checks.
3. Run LLM review in focused passes:
   - section/chunk review
   - full-document consistency review
   - all-table batch review
   - all-image batch review when a vision model is supplied
4. Validate findings:
   - reject self-invalidating results such as "correct", "acceptable", "consistent"
   - require evidence/location/details
   - normalize severity and category
5. Generate Excel:
   - Error Details
   - Summary
   - High/Medium/Low Severity
   - one sheet per category

