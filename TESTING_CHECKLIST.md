# DOC-AI — Testing Checklist
> Everything you need to do, in order, to get the app running and verify the new features.

---

## 🔴 URGENT — Do First (Before Anything Else)

- [ ] **Revoke your GitHub PAT** that was pasted in chat
  - Go to: https://github.com/settings/tokens
  - Find the token starting with `ghp_Znyqps2jiW00...`, click **Delete**
  - Issue a new one only when you need to push again

---

## STEP 1 — Merge the Branch to Main

On GitHub:
- [ ] Open: https://github.com/MrRobot373/DOC-AI-/compare/main...phase-1
- [ ] Click **Create pull request** → review the 13 commits listed → **Merge**
- [ ] Pull locally:
  ```bash
  git checkout main
  git pull
  ```

---

## STEP 2 — Install Dependencies

**Backend (Python):**
```bash
cd backend
pip install -r requirements.txt
```

Verify with:
```bash
python -m unittest discover -s tests -p "test_*.py"
# Expected: Ran 16 tests ... OK
```

**Frontend (Node):**
```bash
cd frontend
npm install
```

Verify with:
```bash
npm run build
# Expected: ✓ built in ~5s  (size warning is normal, not an error)
```

---

## STEP 3 — Run the App (Local Dev, No Docker)

You need **Ollama running on your machine** with at least one model pulled.

**Install Ollama** (if not already): https://ollama.com/download

**Pull models:**
```bash
ollama pull qwen2.5:7b            # good general text model
ollama pull qwen2.5-vl:7b         # vision model for diagrams (optional)
```

**Verify Ollama is live:**
```bash
curl http://localhost:11434/api/tags
# Should return a JSON list of your models
```

**Start the backend:**
```bash
# From the repo root
cd backend
python app.py
# Should print: Open http://localhost:5000 in your browser
```

**Start the frontend (separate terminal):**
```bash
cd frontend
npm run dev
# Opens on http://localhost:5173
```

Open **http://localhost:5173** in your browser.

---

## STEP 4 — Configure the Settings

1. Click your profile icon (top-right) → **API Configuration**
2. Set **Runtime** toggle to: **Local (private)**
3. Set **Host URL** to: `http://localhost:11434`
4. Leave **API Key** blank (not needed for local Ollama)
5. Click **Test & Save**
   - ✅ Expect: "Connection Successful!" and model dropdowns populate with your pulled models
   - ❌ If error: Ollama isn't running — re-run `ollama serve` and try again
6. **Text Model**: select `qwen2.5:7b` (or whatever you pulled)
7. **Vision Model**: select `qwen2.5-vl:7b` if available (shows "(vision)" label)

---

## STEP 5 — Feature Tests (Go Through Each One)

### ✅ TEST 1: Basic Review Works End-to-End
- [ ] Click **Word Document** tab
- [ ] Upload one of the test files from `test_data/` (e.g. `Ultrasmall_Ph4_HardwareDesignDocument_08_04_2026.docx`)
- [ ] Set review mode to **Pro**
- [ ] Click **Start AI Review**
- [ ] Watch the progress bar advance — it should reach 100% within 2–10 minutes
- [ ] Expect: findings list appears, **Download Report (.xlsx)** button appears
- **Pass:** You see real findings with Category, Severity, Comment, Evidence columns

### ✅ TEST 2: Report is Consistent (Determinism Check)
- [ ] Run the **same document a second time** with the same settings
- [ ] Download both Excel reports and compare
- **Pass:** Same findings in both reports (same order, same text). This was broken before — if they're identical, Phase 0 determinism fix works.

### ✅ TEST 3: No Silent AI Failures
- [ ] Temporarily set an **invalid model name** in settings (type `not-a-real-model`)
- [ ] Run a review
- **Pass:** An **amber warning banner** appears at the top of findings saying "N AI passes failed — check model name and API key". Previously this produced a thin report with no warning.
- [ ] Fix the model name back afterward

### ✅ TEST 4: Click-to-Locate Viewer
- [ ] After a completed review, you should see a **two-pane layout**: findings on the left, document preview on the right
- [ ] Click any finding card
- **Pass:** The document pane scrolls to and **highlights** the exact text mentioned in the finding (yellow highlight). Clicking a different finding moves to a different location.
- [ ] If the viewer doesn't show: make sure you're on a Word doc review (not Excel), and findings are present

### ✅ TEST 5: Evidence Column in Excel Report
- [ ] Download the Excel report after a review
- [ ] Open in Excel/LibreOffice
- **Pass:** Report has columns: **No, Page, Section, Evidence, Comment, Fix, Category, Severity, Confidence, Fix Type, Date, Status**
- The **Evidence** column shows the exact text quoted from the document
- The **Confidence** column shows a score like `0.80`

### ✅ TEST 6: Confidence Badges in UI
- [ ] In the findings list, each finding card should show a small grey `XX%` badge (e.g. `80%`)
- **Pass:** Badge is visible. Findings are sorted highest-severity + highest-confidence first.

### ✅ TEST 7: Max Mode Works (Same Engine)
- [ ] Run the same document with **Max** mode
- **Pass:** Review completes, uses the same Excel report format (NOT a different format). Max mode should have fewer findings than Pro (it drops low-confidence ones).

### ✅ TEST 8: Excel File Review Works
- [ ] Switch to **Excel Sheet** tab
- [ ] Upload one of the XLSX files from `test_data/` (e.g. `TICO-ULTRASMALL-PH3-CONCEPT_HSIS_27_02_26 (1).xlsx`)
- [ ] Run in **Pro** mode
- **Pass:** Review completes, findings reference sheet names and row numbers

### ✅ TEST 9: Accuracy Benchmark (Optional but Recommended)
After getting findings from TEST 1, export findings from the UI to JSON (or note the report filename), then run:
```bash
cd backend
python -m eval.score \
  --gold "../test_data/Ultrasmall_Projects/Doc_Review_UltraSmall_HDD (1).xlsx" \
  --doc "../test_data/Ultrasmall_Projects/Ultrasmall_Ph4_HardwareDesignDocument_08_04_2026.docx" \
  --local-only
```
**Pass:** Prints `precision=X  recall=X  f1=X` without crashing. The numbers give you a baseline to track improvement.

---

## STEP 6 — (Optional) Docker Compose — Full Private Stack

Requires **Docker Desktop** installed: https://www.docker.com/products/docker-desktop/

```bash
# From the repo root:
docker compose up --build
# First run takes 5-10 min to build images
```

Once running, pull models INTO the container:
```bash
docker compose exec ollama ollama pull qwen2.5:7b
docker compose exec ollama ollama pull qwen2.5-vl:7b
```

Open **http://localhost:8080** in browser.

In Settings:
- Runtime: **Local**
- Host URL: `http://ollama:11434`  ← use the service name, not localhost
- Click Test & Save

**Pass:** Full review runs on local Ollama; page numbers should be more accurate (Gotenberg handles the PDF rendering automatically).

---

## STEP 7 — Known Gaps (Not Built Yet, Don't Test)

These are NOT done and will not work yet:
- ❌ **Login / Supabase Auth** — the Supabase tables aren't set up for the new schema yet; login may or may not work depending on your existing Supabase config
- ❌ **Review history** — not yet using the new reviews table
- ❌ **Tracked-changes auto-fix** — the current auto-fix is the old silent version; the Word-comment version isn't built yet
- ❌ **PDF.js viewer** (Option B) — Viewer A (docx-preview) is what's built; PDF viewer is Phase 3 remaining

---

## Quick Troubleshooting

| Symptom | Fix |
|---|---|
| Backend won't start (ImportError) | `pip install -r backend/requirements.txt` |
| "No models returned" in settings | Make sure `ollama serve` is running; run `curl http://localhost:11434/api/tags` to confirm |
| Amber warning banner appears | A model pass failed — check the model name is exactly as shown by `ollama list` |
| Viewer pane not visible | Needs a Word doc review with at least 1 finding; Excel shows a table view instead |
| Report has wrong page numbers | Normal for local-dev without LibreOffice/Gotenberg; accurate in Docker Compose stack |
| Frontend 404 on React routes | Use `npm run dev` not opening index.html directly |
