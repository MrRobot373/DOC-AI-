"""
DOC-AI command-line interface — for CI/CD pipelines and scripted reviews.

Reuses the same engine as the web app (parse -> review -> report) with no Flask,
no Supabase. Designed to be dropped into a build pipeline:

    docai review SDD_v3.docx --mode pro --model qwen2.5:32b --fail-on critical
        exit 0  -> clean (or no findings at/above the threshold)
        exit 1  -> findings at/above --fail-on severity exist (blocks the release)
        exit 2  -> execution error

    docai review doc.docx --json                 # findings to stdout for CI parsing
    docai review doc.docx --output report.xlsx   # write the Excel report
    docai eval doc.docx human_report.xlsx         # accuracy score vs a gold report

Run as:  python -m cli ...   (or `docai ...` once installed via the console entry point)
"""

from __future__ import annotations

import argparse
import json
import os
import sys

# Allow running both as `python -m cli` and `python cli.py`
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

SEVERITY_RANK = {"MINOR": 1, "MAJOR": 2, "CRITICAL": 3}


def _build_client(api_key, host):
    from review_engine import create_ollama_client, create_failover_client
    keys = [k.strip() for k in (api_key or "").split(",") if k.strip()]
    if len(keys) > 1:
        return create_failover_client(keys, host)
    return create_ollama_client(keys[0] if keys else "", host)


def cmd_review(args) -> int:
    from doc_parser import parse_document, parse_excel
    from review_engine import review_document
    from report_generator import generate_excel_report

    path = args.document
    if not os.path.exists(path):
        print(f"error: file not found: {path}", file=sys.stderr)
        return 2

    is_excel = path.lower().endswith((".xlsx", ".xls"))
    try:
        parsed = parse_excel(path) if is_excel else parse_document(path)
        if not is_excel:
            try:
                from page_locator import enrich_pages
                enrich_pages(parsed, path)
            except Exception:
                pass

        client = _build_client(args.api_key or os.environ.get("OLLAMA_API_KEY", ""), args.host)

        def progress(msg, pct=None):
            if not args.quiet:
                print(f"  … {msg}", file=sys.stderr)

        status = {}
        findings = review_document(
            client, args.model, parsed,
            progress_callback=progress, review_mode=args.mode,
            vision_model=args.vision_model or None, status_out=status,
        )
    except Exception as e:
        print(f"error: review failed: {e}", file=sys.stderr)
        return 2

    # Output report
    if args.output:
        try:
            generate_excel_report(findings, os.path.basename(path), args.output)
            if not args.quiet:
                print(f"report written: {args.output}", file=sys.stderr)
        except Exception as e:
            print(f"warning: could not write report: {e}", file=sys.stderr)

    # Severity counts
    counts = {"CRITICAL": 0, "MAJOR": 0, "MINOR": 0}
    for f in findings:
        counts[f.get("severity", "MINOR")] = counts.get(f.get("severity", "MINOR"), 0) + 1

    if args.json:
        json.dump({"findings": findings, "counts": counts, "engine_status": status},
                  sys.stdout, ensure_ascii=False, indent=2)
        print()
    elif not args.quiet:
        print(f"\n{len(findings)} findings — "
              f"CRITICAL: {counts['CRITICAL']}, MAJOR: {counts['MAJOR']}, MINOR: {counts['MINOR']}",
              file=sys.stderr)
        if status.get("warning"):
            print(f"warning: {status['warning']}", file=sys.stderr)

    # Exit code based on --fail-on threshold
    if args.fail_on:
        threshold = SEVERITY_RANK.get(args.fail_on.upper(), 99)
        worst = max((SEVERITY_RANK.get(f.get("severity", "MINOR"), 0) for f in findings), default=0)
        if worst >= threshold:
            if not args.quiet:
                print(f"FAIL: findings at/above {args.fail_on.upper()} threshold", file=sys.stderr)
            return 1
    return 0


def cmd_eval(args) -> int:
    """Score local-only (or saved) findings against a human report."""
    try:
        from eval.score import load_gold, score, _findings_from_doc_local_only
        gold = load_gold(args.gold)
        if args.findings:
            with open(args.findings, encoding="utf-8") as fh:
                findings = json.load(fh)
        else:
            findings = _findings_from_doc_local_only(args.document)
        result = score(findings, gold)
        json.dump(result, sys.stdout, indent=2)
        print()
        return 0
    except Exception as e:
        print(f"error: eval failed: {e}", file=sys.stderr)
        return 2


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="docai", description="DOC-AI technical document reviewer (CLI)")
    sub = p.add_subparsers(dest="command", required=True)

    r = sub.add_parser("review", help="Review a document and optionally fail on severity")
    r.add_argument("document", help="Path to .docx / .xlsx")
    r.add_argument("--mode", choices=["normal", "pro", "max"], default="pro")
    r.add_argument("--model", default=os.environ.get("DOCAI_MODEL", "qwen2.5:7b"))
    r.add_argument("--vision-model", default=os.environ.get("DOCAI_VISION_MODEL", ""))
    r.add_argument("--host", default=os.environ.get("OLLAMA_HOST", "http://localhost:11434"))
    r.add_argument("--api-key", default="")
    r.add_argument("--output", help="Write the Excel report to this path")
    r.add_argument("--json", action="store_true", help="Emit findings as JSON to stdout")
    r.add_argument("--fail-on", choices=["critical", "major", "minor"], help="Exit 1 if findings at/above this severity")
    r.add_argument("--quiet", action="store_true")
    r.set_defaults(func=cmd_review)

    e = sub.add_parser("eval", help="Score findings against a human-reviewed report")
    e.add_argument("document", help="Path to the reviewed document")
    e.add_argument("gold", help="Path to the human .xlsx report (ground truth)")
    e.add_argument("--findings", help="Optional findings JSON to score (else local-only checks)")
    e.set_defaults(func=cmd_eval)
    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
