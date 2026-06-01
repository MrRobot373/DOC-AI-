# DOC-AI — Running & Deploying

Two supported targets: **on-prem (Docker Compose, private/local LLM)** and **cloud SaaS**.

---

## A. On-prem / local — Docker Compose (recommended for confidential docs)

Everything runs on your machine/network; documents never leave it.

```bash
docker compose up --build
```

This starts four services:

| Service | Port | Purpose |
|---------|------|---------|
| `ollama` | 11434 | Local open-source LLM runtime |
| `gotenberg` | 3000 | DOCX → PDF rendering (accurate page numbers) |
| `backend` | 10000 | Flask review API |
| `frontend` | 8080 | React UI |

**First-time setup — pull models into Ollama:**

```bash
docker compose exec ollama ollama pull qwen2.5:7b        # text/reasoning
docker compose exec ollama ollama pull qwen2.5-vl:7b     # vision (diagrams/graphs)
```

Then open **http://localhost:8080** → profile → **API Configuration**:
- **Runtime:** Local
- **Host URL:** `http://ollama:11434`  (the backend reaches Ollama by service name)
- **API Key:** leave blank (not needed for local)
- Click **Test & Save** — the model dropdowns populate from your local Ollama.

Real page numbers work automatically because `GOTENBERG_URL` is wired to the
`gotenberg` service. Without it the tool falls back to heuristic pages.

### Useful env (set in `docker-compose.yml` or a `.env` file)
- `ALLOWED_ORIGINS` — comma-separated allowed origins (default locks to the frontend).
- `MAX_UPLOAD_MB` — max upload size (default 50).
- `GOTENBERG_URL` — PDF renderer (default `http://gotenberg:3000`).
- `VITE_SUPABASE_URL` / `VITE_SUPABASE_ANON_KEY` — optional; only if you want
  Supabase login/history. Pure on-prem can leave them blank.

---

## B. Cloud SaaS (current hosted setup)

- **Frontend:** Vercel (build with `VITE_API_URL`, `VITE_SUPABASE_*`).
- **Backend:** Render (`render.yaml`) — `gunicorn -c gunicorn.conf.py app:app`.
- **Auth/persistence:** Supabase.
- **LLM:** Ollama Cloud (Runtime = Cloud, host `https://ollama.com`, API key required).
- Set `GOTENBERG_URL` to a reachable Gotenberg instance for accurate pages, or
  leave unset to use heuristic pages.

---

## C. Local dev (no Docker)

```bash
# Backend
cd backend
pip install -r requirements.txt
python app.py                      # http://localhost:5000

# Frontend (separate terminal)
cd frontend
npm install
npm run dev                        # http://localhost:5173
```

Point the UI at a local Ollama (`http://localhost:11434`, Runtime = Local) or
Ollama Cloud. For accurate pages, run Gotenberg and set `GOTENBERG_URL`, or
install LibreOffice (`soffice`) — otherwise pages use the heuristic fallback.

---

## Tests

```bash
cd backend
python -m unittest discover -s tests -p "test_*.py"

# Accuracy vs a human-reviewed report:
python -m eval.score --gold ../test_data/<human_report>.xlsx --doc <doc>.docx --local-only
```
