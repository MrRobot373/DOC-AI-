# DOC-AI — Production Analysis & Roadmap

> Deep audit of the DOC-AI technical-document review tool, the root causes of
> "bad / inconsistent reports", and a phased plan to make it a production-grade
> analyzer that runs on **open-source LLMs via Ollama (local-first, cloud-optional)**.
>
> Status: **analysis + plan only — no code changed yet.** Work the checklists top-down.

---

## 0. Executive summary

The tool is architecturally reasonable (deterministic local checks + multi-pass LLM +
vision + Excel report), but it *feels* unfinished because of four concrete, fixable issues:

1. **The LLM passes silently fail on most runs** → reports collapse to noisy local-only checks.
2. **No determinism controls** on LLM calls → same doc, different report every run.
3. **Page numbers are guessed and usually wrong** → the whole report looks untrustworthy.
4. **Pro mode has no hallucination guard** → invented findings + low-value noise bury the real issues.

Fixing items 1–4 (Phase 0 + the page work) is ~80% of the perceived quality problem.
Everything else (auth, durable queue, local-Ollama default, unified engine) is what makes
it *sellable to enterprise clients like Wise*.

**Decisions locked in for this plan:**
- LLM target: **both local + cloud, user-selectable** (local default, cloud opt-in).
- Working mode: **plan-first** — implement only after approval of each phase.

---

## 1. Root-cause diagnosis — why reports are bad & inconsistent

### RC1 — LLM passes fail silently, leaving only local checks
The dashboard ships a **hardcoded model list** containing models that do not exist on Ollama:
`frontend/src/pages/Dashboard.tsx:43-52` →
`nemotron-3-super:cloud`, `qwen3.5:397b-cloud`, `minimax-m2.7:cloud`, `kimi-k2.5:cloud`.

When `client.chat` is handed a non-existent model it raises, and nearly every LLM pass
swallows the error:
- `backend/review_engine.py:286-287` (table pass) → `except Exception: pass`
- `backend/review_engine.py:306-308` (image pass) → `except Exception: pass`
- `backend/review_engine.py:901-909` (chunk pass) → logs + returns `[]`
- `backend/review_engine.py:958-973` (consistency) → returns a single cryptic error finding

**Result:** a "bad" run silently degrades to *local regex checks only*; a "good" run
(valid model + quota) returns the full multi-pass report. Same input, very different output.

> The backend already exposes `GET/POST /api/models` (`backend/app.py:152-163`) that returns
> the **real** models from `client.list()`. The dashboard never calls it.

### RC2 — No determinism on LLM calls
All four call sites use `temperature: 0.05` but **no `seed` and no structured-output `format`**:
- `backend/review_engine.py:893-897` (chunk)
- `backend/review_engine.py:950-954` (consistency)
- `backend/review_engine.py:1027-1031` (tables)
- `backend/review_engine.py:1115-1122` (images)

Ollama supports `options={"seed": N, "temperature": 0}` and `format=<json-schema>` for
guaranteed-parseable output. Without these you get run-to-run variance *and* rely on a
brittle regex JSON fallback (`backend/review_engine.py:1201-1235`).

### RC3 — Page numbers are heuristic guesses
`backend/doc_parser.py:114-159` estimates pages from `len(text)//80` lines-per-page math,
calibrated only at heading boundaries via the TOC (`backend/doc_parser.py:200-216`).
`PLAN.md` documents this as "inflated 40..98 drift" and specifies a Gotenberg/LibreOffice
PDF-render fix — **which was never implemented** (no `backend/page_locator.py` exists).
A finding citing "Page 43" for content on page 4 destroys trust in the whole report.

### RC4 — Pro mode has no evidence grounding; noise buries signal
Max mode has a good filter that drops any LLM finding whose evidence tokens are absent from
the document (`kimi_style_analyzer/analyze.py:498-530`). **Pro mode has nothing equivalent**,
so hallucinations pass through. Simultaneously, low-value local checks generate MINOR noise
(double spaces, `5V` spacing, unmatched brackets) that drowns the CRITICAL findings.

---

## 2. Critical bugs (correctness / security)

| ID | Location | Problem | Fix |
|----|----------|---------|-----|
| B1 | `backend/app.py:504-506` | **Path traversal on upload** — `file.filename` concatenated into save path unsanitized. | `werkzeug.utils.secure_filename`. |
| B2 | `backend/app.py:625-667` | **Path traversal on download** — `report_filename`/`filename` go straight into `os.path.join`. | Validate against an allowlist / `os.path.basename` + realpath check. |
| B3 | `backend/app.py:39`, `:535-557` | **No auth, open CORS, guessable 8-hex IDs** — `/api/progress/<id>` leaks other tenants' findings. | JWT gate + per-`user_id` scoping; restrict CORS origins. |
| B4 | `backend/app.py:82-130`, `:340-353` | **Lost-update race** — every progress tick reloads + rewrites the *entire* store (and re-upserts all of it to Supabase); concurrent reviews clobber each other; restarts lose in-flight work. | Per-review row + atomic update / real queue. |
| B5 | `backend/doc_parser.py:35-40` | **Fallback parse returns malformed dict** (missing `headings/tables/toc/raw_text`, uses `title` not `heading`) → `get_document_summary` / `_section_to_text` `KeyError`. | Return a complete, schema-valid skeleton. |
| B6 | `backend/review_engine.py:1595-1611` | **Unmatched-bracket check is logically broken** (operator precedence + per-paragraph scope) → false positives on normal prose. | Rework or remove. |
| B7 | `backend/review_engine.py:297` | **Vision gating by model-name substring** (`vl/vision/llava/qwen`) — text-only `qwen3.5` passes; real vision models without those tokens are skipped. | Explicit `is_vision` capability flag. |
| B8 | `frontend/src/pages/Dashboard.tsx:173` | Default model fallback `gpt-oss:120b-cloud` matches no list entry → another "model not found" path. | Derive default from fetched model list. |
| B9 | `backend/doc_fixer.py:204-209` | Case-insensitive find + replace can replace the **wrong case / wrong occurrence**, risking client-doc corruption. | Tracked changes / comments instead of silent edits. |

---

## 3. Dead, bad, and overfit logic

- **Overfit to one TICO document** — hardcoded strings/regexes that do nothing for a generic
  client doc and look unprofessional:
  - `kimi_style_analyzer/analyze.py:168-181` — typos `"turns ration"`, `"Voit"`.
  - `kimi_style_analyzer/analyze.py:236-270` — `650V.*650V`, `1500W.*1350W`, `"custom part from TICO 1500W"`.
  - → Replace with a **per-project glossary / rules file** (JSON), not baked-in strings.
- **False-positive factories** (local checks):
  - `backend/review_engine.py:1661-1680` `_check_unit_standardization` flags `5V`/`3V3` — usually **net/signal names** in HW docs, not measurements.
  - `backend/review_engine.py:415-470` `_check_font_consistency` flags Symbol/Wingdings/monospace fonts used legitimately for Ω, µ, checkboxes, code.
  - `backend/review_engine.py:543-650` `_check_cross_references` only recognizes captions like `Figure 5:`/`Table 3 —`; Word SEQ-field / *Caption*-style figures are invisible → real figures reported as "broken references".
- **Dead infrastructure** — `USE_THREADING = True` is hardcoded (`backend/app.py:51`), so the
  entire Huey/Redis/SqliteHuey branch (`backend/app.py:53-66`, `:526-530`) and `start_worker.bat`
  are dead, yet `redis`/`huey` are still installed.
- **Unused dependency** — `nltk` is in `backend/requirements.txt` and `build.sh` downloads
  `punkt`/`punkt_tab` on every deploy, but nltk is never imported. Pure deploy cost + failure surface.
- **Unpinned requirements** — `backend/requirements.txt` has zero version pins → a deploy can
  pull a breaking `python-docx`/`ollama`/`supabase` and change behavior silently. This is itself
  a source of "it worked yesterday" inconsistency.
- **Two parallel engines + two report formats** — Normal/Pro (`review_engine.py`) vs Max
  (`kimi_style_analyzer/analyze.py`) duplicate ~70% of logic. Max findings lose page info
  (`page: "-"`, `backend/app.py:268`) and regenerate-on-download skips max mode (`backend/app.py:639`).

---

## 4. Parsing & analysis quality gaps

- **Tables reviewed in isolation** — table data never appears in the same LLM context as the
  prose, so "text says 3.3 V but table says 5 V" is structurally hard to catch.
- **Char-based chunking, no overlap** — `get_section_chunks(max_chars=5000)`
  (`backend/doc_parser.py:849`) splits by characters not tokens, with no boundary overlap;
  issues spanning a chunk edge are missed, and an oversized section/table can blow the model's context.
- **Merged cells duplicate content** — python-docx returns identical text for each cell of a
  merge, inflating decimal/duplication checks.
- **Images carry no location** — extracted from `doc.part.rels` (`backend/doc_parser.py:695-741`)
  with no section/page, so vision findings can only say "Image 3"; header logos are reviewed as content.
- **No content from text boxes / content controls / SmartArt** — `doc.paragraphs` misses these.

---

## 5. The Ollama story — local-first, cloud-optional

The app is currently hardwired to **Ollama Cloud** (`https://ollama.com` + bearer keys) in the
backend defaults and the frontend. For clients like **Wise**, confidential design docs must
**never leave their network**, so local/on-prem must be the default.

**Target design (user-selectable, per decision):**
- A **"Runtime" toggle** in settings: `Local` (default) vs `Cloud`.
  - Local → host `http://localhost:11434`, API key optional/empty.
  - Cloud → host `https://ollama.com`, API key required (existing failover logic reused).
- **Always fetch models from the chosen host** via `/api/models` → users only pick models that exist.
- Recommended open-source models (documented in README):
  - Text/reasoning: `qwen2.5:32b` / `qwen2.5:72b`, `llama3.3:70b`, `gpt-oss`.
  - Vision: `qwen2.5-vl`, `llama3.2-vision`, `minicpm-v`.
- Mark vision-capable models explicitly (fixes B7) rather than sniffing the name.

---

## 6. Production-readiness checklist (enterprise)

- [ ] **Auth + multi-tenancy** — JWT gate on all `/api/*`; scope every review to `user_id`; stop serving findings to unauthenticated pollers.
- [ ] **Durable job queue + `reviews` table** — replace in-process thread + JSON file with Redis/RQ (or Postgres-backed) and one DB row per review (status + findings as JSONB). Fixes B4, survives restarts, enables `workers > 1`.
- [ ] **Encrypt secrets** — API keys are plaintext in Supabase `user_settings`; encrypt at rest or use a secret store. Remove hardcoded personal email (`backend/app.py:166`).
- [ ] **Observability** — structured logging, per-pass timing, and a user-visible banner when an LLM pass fails (no more `except: pass`).
- [ ] **Reproducible builds** — pin all deps, remove nltk, remove dead Huey/Redis path, simplify `build.sh`.
- [ ] **Safe auto-fix** — emit Word tracked-changes/comments instead of silent run edits (B9).
- [ ] **Rate limiting + max upload size** on `/api/review` and `/api/feedback`.

---

## 7. Feature roadmap — toward best-in-class

1. **Evidence-grounded findings everywhere** — port Max-mode token grounding
   (`kimi_style_analyzer/analyze.py:498-530`) into Pro; require every LLM finding to quote an
   exact in-document substring or be dropped. *Biggest single quality lever.*
2. **Real page numbers** — implement the Gotenberg/LibreOffice → PDF anchor step from `PLAN.md`,
   heuristic as fallback, record `metadata.page_source`.
3. **One unified engine + one report** — merge Normal/Pro/Max into a single pipeline with a
   strictness knob; retire the duplicate kimi engine (keep its good ideas).
4. **Deterministic + cached** — `seed` + `format=json-schema` + `temperature:0`, and cache
   `(chunk_hash, model) → findings` so re-runs are identical and free.
5. **Token-aware chunking with overlap** + table content included in prose context.
6. **Confidence score + evidence quote per finding**, sortable in the UI — turns "noisy" into "trusted".
7. **Golden-file regression suite** — fixed docs with known expected findings, run in CI, to
   *measure* precision/recall and prevent regressions. (Today: only `backend/tests/test_toc_validation.py`.)

---

## 8. Accuracy engineering — fixing false positives & false negatives

This is the core "it's accurate but not 100%" problem. There are exactly two failure modes,
and they need different fixes:

- **False positive (FP):** flags something that isn't actually wrong → looks noisy/untrustworthy.
- **False negative (FN):** misses a real issue → looks incomplete.

You cannot reach high accuracy by "tuning prompts and hoping." You need a **measurement loop**.

### 8.1 The single most important move — build a golden set from data you already have
`test_data/` already contains **human-reviewed TICO reports** (e.g.
`*_Review_*.xlsx`, `...MatsSanFeedback.xlsx`, `Ultrasmall-CAE-CFD-Endo-san-comments.xlsx`).
These are **ground truth** — real reviewers' findings on real docs. Use them:

1. For each sample doc, extract the human findings into a labeled list ("true issues").
2. Run the tool, then compute per-category **precision** (of what we flagged, how many were real)
   and **recall** (of the real issues, how many we caught).
3. Every prompt/check change is scored against this set. This is how you move from
   "feels accurate" to "91% precision / 84% recall and rising" — and how you stop regressions.

> Without this, accuracy work is guesswork. With it, every change is measurable.

### 8.2 Reducing false positives (stop flagging wrong things)
| Cause (in current code) | Fix |
|---|---|
| No evidence grounding in Pro mode (`review_engine.py`) | Require every LLM finding to quote an exact in-document substring; **drop it if the quote isn't found** (port Max-mode logic `analyze.py:498-530`). |
| **Self-verification missing** | Add a second "critic" LLM pass: feed each candidate finding + its evidence back and ask "Is this a real error? keep/drop + reason." Cuts FPs sharply for low extra cost. |
| Context-blind regex (`5V` net name flagged as unit error) | Maintain a net/signal-name allowlist; skip identifier-shaped tokens; only flag in measurement context. |
| `_check_cross_references` misses real figures → flags valid refs as broken | Detect SEQ-field / *Caption*-style figures & tables (Phase 1). |
| Symbol/Wingdings/monospace fonts flagged | Whitelist symbol fonts in `_check_font_consistency`. |
| Same issue found by local **and** LLM **and** kimi → looks like noise | One source of truth per check; dedupe across sources by **location + category + evidence overlap**, not fuzzy text ratio. |
| Nondeterminism (`temperature 0.05`, no seed) | `seed` + `temperature:0` + JSON-schema `format` (Phase 0). |
| Everything shown as equal | **Confidence score** per finding + a strictness threshold; hide/sort low-confidence. |

### 8.3 Reducing false negatives (stop missing things)
| Cause | Fix |
|---|---|
| LLM passes silently fail → whole categories missing | Surface failures + retry + validate model (Phase 0, RC1). |
| Char-based chunking, **no overlap** → boundary issues lost | Token-aware chunks with overlap. |
| Each chunk reviewed alone → cross-section terminology/contradiction invisible | A dedicated **document-wide** pass over an entity/terminology map (acronyms, signal names, part numbers, rails). |
| Tables reviewed in isolation from prose | Include relevant table rows in the prose chunk context so "text says 3.3 V, table says 5 V" is catchable. |
| One generic mega-prompt per chunk | **Category-specialized passes** (a focused terminology pass, a numeric/units pass, a cross-ref pass). Focus raises recall per category. |
| Single sample | **Ensemble / multi-sample voting** — run a chunk 2–3× (or with 2 models), union the findings, then ground + dedupe. Grounding kills the extra FPs the union introduces, so recall ↑ without FP ↑. |
| Parser never sees text boxes / content controls / SmartArt / comments | Extend parser coverage. |
| Model lacks project knowledge | Inject a **project glossary / rules file** (acronyms, conventions, requirement IDs) into prompts; optionally RAG over the requirements/SRS doc. |

### 8.4 Reframe the division of labor
- **Deterministic local checks** should be **high-precision only** — flag *only* when certain
  (numbers, structure, exact patterns). Move all fuzzy/judgment work to the LLM.
- **LLM** handles ambiguity, but **must ground every claim in a quote** and **pass the critic**.
- This split is what makes the report both trustworthy (few FPs) and thorough (few FNs).

---

## 9. Side-by-side document viewer + jump-to-issue

Goal: show the technical doc and the findings side by side; clicking a finding scrolls to and
highlights the exact problem location. (Reviewers want to be taken to the issue, not read a list.)

### 9.1 How it works — the anchor is the key
Each finding must carry an **exact evidence quote** + **`para_index`** (parser already emits
`[¶N]` markers, `doc_parser.py:876-889`). On click: search the rendered doc for the quote →
scroll → highlight. This reuses the *same* evidence grounding from §8.2 — quality work and the
viewer are one investment.

### 9.2 Two build options
- **Option A (ship fast, pure browser):** render the uploaded `.docx` → HTML with
  [`docx-preview`](https://www.npmjs.com/package/docx-preview) or `mammoth.js`; click → find
  quote in HTML → scroll + CSS highlight. No server changes. Lower layout fidelity; pages fuzzy.
- **Option B (high fidelity, recommended end state):** server renders `.docx` → **PDF** (the
  same Gotenberg/LibreOffice step added for page numbers in Phase 1); frontend renders with
  **PDF.js** (built-in page nav + text-layer highlight). Looks exactly like Word; reuses page work.
- **Excel:** render parsed rows as an HTML table; click → scroll to sheet + row (indices already parsed).

### 9.3 Jump precision by finding type
| Finding type | Precision |
|---|---|
| Text issue with exact quote | 🟢 Exact highlight |
| Table cell issue | 🟢 Exact (row/col known) |
| Section-level finding | 🟡 Scroll to heading |
| Image / vision finding | 🟡 Scroll to image — needs image-position tracking (parser gap) |
| Whole-document finding | 🔴 No single spot — pin to top |

### 9.4 Data-model change required
Add to each finding: `evidence` (exact quote), `para_index` (already parsed), and for images a
`position` anchor (new). Render + scroll + highlight is standard frontend work.

---

## 10. Auto-fix capability matrix

Auto-fix is possible for the deterministic/basic bucket only; critical engineering issues must
stay human-reviewed. **For client docs, never silently edit** — emit Word **Tracked Changes**
(engineer clicks Accept/Reject) and insert **Word Comments** for non-fixable findings anchored to
the text. (Current `doc_fixer.py` does silent run edits and only handles a tiny slice — see B9.)

| Category | Auto-fixable? | Approach |
|---|---|---|
| Grammar & Spelling (clear typo + exact replacement) | 🟢 Auto | Tracked-change replacement |
| Repeated words | 🟢 Auto | Delete duplicate |
| Formatting: double spaces, trailing space, empty paras | 🟢 Auto | Normalize |
| Subscript/Superscript (`<sub>` tags, `^` carets) | 🟢 Auto | Convert to real Word formatting |
| Terminology consistency (VDD vs VCC) | 🟡 Semi (confirm canonical term, then global) | Tracked-change replace-all |
| Units spacing/casing (`5V`→`5 V`) | 🟡 Semi (risk: net names) | Confirm per token |
| Decimal-place standardization (table column) | 🟡 Semi (pick N) | Pad/round in cells |
| Date format normalization | 🟡 Semi (pick target format) | Replace |
| Cross-reference errors | 🔴 Manual | Comment only — tool can't know correct target |
| Logical/technical contradictions, power/voltage margins | 🔴 Manual | Comment only |
| Schematic/vision findings (shorted pins, wrong symbol) | 🔴 Manual | Comment only |
| Missing test criteria / design elements | 🔴 Manual | Comment only |

**Realistic yield:** of ~100 findings, ~30–50 (grammar/spelling/formatting/terminology) are
auto/semi-fixable; the critical engineering findings stay human-reviewed (correctly).

---

## 11. Phased implementation plan (work top-down)

### Phase 0 — Reliability & consistency (do first; highest ROI, lowest risk)
- [ ] Wire dashboard to `GET /api/models` from the selected host; **delete the hardcoded model list** (`Dashboard.tsx:43-52`). Default model = first returned (B8).
- [ ] Add `seed` + `temperature:0` + `format` (JSON schema) to all 4 LLM call sites (RC2).
- [ ] Replace `except: pass` in table/image/chunk passes with error capture; surface a
      "N AI passes failed" banner + per-pass status to the user (RC1).
- [ ] Port evidence grounding into Pro mode; drop findings without an in-document quote (RC4).
- [ ] Mute/merge noisy low-value checks (unit spacing on net names, font Symbol/Wingdings,
      broken bracket check) behind the strictness setting (RC4, B6).
- [ ] Fix fallback-parse schema so odd files don't crash the review (B5).
- [ ] Tag every finding with `evidence` (exact quote) + `para_index` — unblocks both the
      critic pass and the side-by-side viewer (§8, §9).
- [ ] Add the **self-verification critic pass** to cut false positives (§8.2).

### Phase 1 — Trust the output
- [ ] Implement `backend/page_locator.py` (Gotenberg/LibreOffice PDF anchoring) per `PLAN.md`; wire into `app.py` after parse.
- [ ] Rework `_check_cross_references` to recognize Word SEQ-field / Caption-style figures & tables.
- [ ] Merge the two engines + two report formats into one; preserve evidence + page columns.
- [ ] **Build the golden-set accuracy harness** from `test_data/` human reports; report
      per-category precision/recall (§8.1).
- [ ] **Side-by-side viewer (Option A)** — `docx-preview` + click-to-jump/highlight (§9.2).
- [ ] Make deterministic checks high-precision; move fuzzy work to the LLM (§8.4).

### Phase 2 — Enterprise hardening
- [ ] Local-vs-cloud runtime toggle (per decision); local default; fetch models per host.
- [ ] JWT auth + per-tenant scoping; restrict CORS; secure_filename + download path validation (B1, B2, B3).
- [ ] Durable queue + `reviews` table; atomic per-review updates (B4); enable multi-worker.
- [ ] Encrypt stored API keys; remove personal email; pin deps; drop nltk + dead Huey/Redis.

### Phase 3 — Best-in-class
- [ ] Confidence scoring + evidence display + sort; strictness threshold (§8.2).
- [ ] Finding cache by content hash.
- [ ] Token-aware chunking with overlap; tables in prose context (§8.3).
- [ ] Category-specialized passes + ensemble/multi-sample voting for recall (§8.3).
- [ ] Project glossary / rules file injected into prompts (§8.3).
- [ ] Golden-file precision/recall CI suite (gate on regressions).
- [ ] **Side-by-side viewer (Option B)** — PDF.js over the rendered PDF (§9.2).
- [ ] Tracked-changes + comment-based auto-fix per the §10 matrix (B9).

---

## 12. Quick reference — key files

| File | Role | Notable lines |
|------|------|---------------|
| `backend/app.py` | Flask API, job orchestration, state store | 51 (dead threading flag), 82-130 (race), 504-506 (B1), 625-667 (B2) |
| `backend/review_engine.py` | Normal/Pro engine: 15 local checks + LLM passes | 286-308 (silent fails), 543-650 (xref FP), 893-897 (no seed), 1238-1252 (O(n²) dedup) |
| `backend/doc_parser.py` | DOCX/XLSX → structured doc | 35-40 (B5), 114-159 (page heuristic) |
| `backend/report_generator.py` | Pro Excel report | — |
| `backend/doc_fixer.py` | Auto-fix (silent run edits) | 204-209 (B9) |
| `kimi_style_analyzer/analyze.py` | Max engine (duplicate) | 168-270 (overfit), 498-530 (good grounding) |
| `frontend/src/pages/Dashboard.tsx` | Upload UI, settings, findings | 43-52 (fake models), 173 (bad default) |
| `PLAN.md` | Spec for unimplemented page-number fix | — |
