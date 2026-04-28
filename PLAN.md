# Open-Source Page Number Fix via Gotenberg

## Summary
- Make PDF-rendered pages the primary source of truth for Word page numbers.
- Restore a stable heuristic fallback in the parser so reviews still run when the renderer is unavailable.
- Keep scope limited to wrong page numbers only; do not change `"-"` page placeholders for image/table findings in this pass.

## Key Changes
- In [backend/doc_parser.py](c:/Users/yash%20badgujar/Downloads/TICO/backend/doc_parser.py:191), remove the current `id(body_child)` / `id(para._element)` page-map approach from the uncommitted rewrite.
- Revert the fallback parser behavior to the last committed TOC-calibrated baseline, or replace the bad lookup with a stable element-keyed/object-keyed map only for fallback mode. Do not keep the current inflated body-order estimator as the primary path.
- Add a new renderer-backed page enrichment step, preferably in a new module such as `backend/page_locator.py`.
- The renderer step should:
  - POST the uploaded `.docx` to Gotenberg’s LibreOffice conversion endpoint.
  - Set `updateIndexes=true` so TOC and index fields are refreshed during conversion.
  - Read the returned PDF with `pypdf`.
  - Build normalized text for each PDF page.
  - Reassign page numbers for headings and paragraphs by monotonic text anchoring:
    - headings: exact normalized heading text match first
    - paragraphs: normalized leading-token snippet match, searched from the last assigned page forward
    - fallback: use the enclosing section start page if a paragraph snippet is not found
  - Write the resolved page back into `parsed["headings"]`, each section’s `page`, each paragraph’s `page`, and `parsed["statistics"]["total_pages"]`.
  - Record `parsed["metadata"]["page_source"]` as `"gotenberg"` or `"heuristic"`.
- Wire the enrichment into [backend/app.py](c:/Users/yash%20badgujar/Downloads/TICO/backend/app.py:270) immediately after `parse_document(filepath)` and before `review_document(...)`.
- Add config:
  - `GOTENBERG_URL` required for renderer mode
  - default local assumption: `http://localhost:3000`
  - if unset or conversion fails, continue with heuristic fallback and log the reason

## Public Interfaces / Types
- New env var: `GOTENBERG_URL`
- New parsed-doc metadata field: `metadata.page_source` with values `"gotenberg"` or `"heuristic"`
- No frontend API contract changes required; findings should continue consuming existing `page` fields

## Test Plan
- Parser regression test: confirm the sample DOCX no longer inflates early headings, e.g. `High-level requirements` resolves to page `4` instead of the current broken `43`.
- Renderer integration test: when Gotenberg is available, assert that late-document headings land on PDF-derived pages and that `total_pages` matches the converted PDF.
- Monotonicity test: paragraph and heading pages never decrease as document order advances.
- Fallback test: with `GOTENBERG_URL` missing or the service unavailable, parsing still completes and `metadata.page_source == "heuristic"`.
- Acceptance check on the provided sample file: report pages should stay in a realistic range tied to the rendered PDF, not the current inflated `40..98` drift.

## Assumptions
- Primary open-source renderer is Gotenberg, not a host-installed LibreOffice binary.
- `pypdf` is acceptable as the PDF text extraction dependency.
- This pass does not attempt to fill `"-"` pages for image/table findings.
- The current uncommitted body-order estimator is not retained as the main page-number source.
