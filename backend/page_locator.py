"""
Page locator — replace heuristic page numbers with PDF-accurate ones.

The DOCX is rendered to PDF (Gotenberg preferred, headless LibreOffice fallback),
then each heading/paragraph is anchored to the PDF page whose text contains it
via a monotonic forward search. This fixes the long-standing "wrong page numbers"
problem (see PLAN.md). If no renderer is available, the existing heuristic pages
are left untouched and metadata.page_source is set to "heuristic".

Public API:
    enrich_pages(parsed, docx_path, gotenberg_url=None) -> parsed (mutated)
    assign_pages(parsed, page_texts) -> parsed   # pure, unit-testable
"""

from __future__ import annotations

import io
import os
import re
import shutil
import subprocess
import tempfile


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip().lower()


# ------------------------------------------------------------------
# Rendering
# ------------------------------------------------------------------
def _render_with_gotenberg(docx_path: str, gotenberg_url: str, timeout: int):
    import requests  # local import so the module loads without requests at import time

    url = gotenberg_url.rstrip("/") + "/forms/libreoffice/convert"
    with open(docx_path, "rb") as fh:
        files = {"files": (os.path.basename(docx_path), fh)}
        # updateIndexes=true refreshes TOC/index fields during conversion.
        data = {"updateIndexes": "true"}
        resp = requests.post(url, files=files, data=data, timeout=timeout)
    if resp.status_code == 200 and resp.content:
        return resp.content
    raise RuntimeError(f"Gotenberg returned {resp.status_code}")


def _render_with_soffice(docx_path: str, timeout: int):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, docx_path],
            check=True, timeout=timeout,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        pdfs = [f for f in os.listdir(tmp) if f.lower().endswith(".pdf")]
        if not pdfs:
            return None
        with open(os.path.join(tmp, pdfs[0]), "rb") as fh:
            return fh.read()


def _render_pdf(docx_path: str, gotenberg_url, timeout: int):
    """Return (pdf_bytes, source) or (None, 'heuristic')."""
    if gotenberg_url:
        try:
            return _render_with_gotenberg(docx_path, gotenberg_url, timeout), "gotenberg"
        except Exception as e:
            print(f"[page_locator] Gotenberg render failed: {e}")
    try:
        pdf = _render_with_soffice(docx_path, timeout)
        if pdf:
            return pdf, "soffice"
    except Exception as e:
        print(f"[page_locator] soffice render failed: {e}")
    return None, "heuristic"


def _pdf_page_texts(pdf_bytes: bytes) -> list:
    try:
        from pypdf import PdfReader
    except ImportError:
        print("[page_locator] pypdf not installed; skipping PDF page extraction")
        return []
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        return [_normalize(page.extract_text() or "") for page in reader.pages]
    except Exception as e:
        print(f"[page_locator] PDF text extraction failed: {e}")
        return []


# ------------------------------------------------------------------
# Anchoring (pure — unit-testable without a real renderer)
# ------------------------------------------------------------------
def _snippet(text: str, words: int = 6, max_chars: int = 40) -> str:
    norm = _normalize(text)
    if not norm:
        return ""
    return " ".join(norm.split()[:words])[:max_chars]


def assign_pages(parsed: dict, page_texts: list) -> dict:
    """
    Anchor every paragraph/heading to a PDF page by monotonic forward search.

    `page_texts` is a list of normalized page strings (index 0 == page 1).
    Pages never decrease as document order advances (monotonic), which matches
    how a document actually paginates and avoids the old heuristic's drift.
    """
    if not page_texts:
        return parsed

    total_pages = len(page_texts)
    cursor = 0  # 0-indexed page we are currently on
    index_to_page = {}

    for section in parsed.get("sections", []):
        section_page = None
        for para in section.get("paragraphs", []):
            snippet = _snippet(para.get("text", ""))
            if snippet:
                found = None
                for offset in range(cursor, total_pages):
                    if snippet in page_texts[offset]:
                        found = offset
                        break
                if found is not None:
                    cursor = found
            page = cursor + 1  # 1-indexed
            para["page"] = page
            if para.get("index") is not None:
                index_to_page[para["index"]] = page
            if section_page is None:
                section_page = page
        if section_page is not None:
            section["page"] = section_page

    # Update the flat headings list via their paragraph index.
    for heading in parsed.get("headings", []):
        idx = heading.get("index")
        if idx in index_to_page:
            heading["page"] = index_to_page[idx]

    parsed.setdefault("statistics", {})["total_pages"] = total_pages
    return parsed


def enrich_pages(parsed: dict, docx_path: str, gotenberg_url=None,
                  timeout: int = 120, pdf_save_dir: str = None) -> dict:
    """
    Render the DOCX to PDF and reassign page numbers from the rendered pages.
    Falls back to the existing heuristic pages when no renderer is available.
    Records parsed['metadata']['page_source'] = gotenberg|soffice|heuristic.

    If `pdf_save_dir` is given and rendering succeeds, saves the PDF there so
    it can later be served to the browser for the PDF.js viewer. The saved
    filename is stored in parsed['metadata']['pdf_filename'].
    """
    parsed.setdefault("metadata", {})
    gotenberg_url = gotenberg_url or os.environ.get("GOTENBERG_URL")

    pdf_bytes, source = _render_pdf(docx_path, gotenberg_url, timeout)
    if not pdf_bytes:
        parsed["metadata"]["page_source"] = "heuristic"
        return parsed

    page_texts = _pdf_page_texts(pdf_bytes)
    if not page_texts:
        parsed["metadata"]["page_source"] = "heuristic"
        return parsed

    assign_pages(parsed, page_texts)
    parsed["metadata"]["page_source"] = source

    # Persist the PDF so the viewer can serve it.
    if pdf_save_dir and pdf_bytes:
        try:
            os.makedirs(pdf_save_dir, exist_ok=True)
            base = os.path.splitext(os.path.basename(docx_path))[0]
            pdf_name = f"{base}_rendered.pdf"
            pdf_path = os.path.join(pdf_save_dir, pdf_name)
            with open(pdf_path, "wb") as fh:
                fh.write(pdf_bytes)
            parsed["metadata"]["pdf_filename"] = pdf_name
        except Exception as e:
            print(f"[page_locator] PDF save failed: {e}")

    return parsed
