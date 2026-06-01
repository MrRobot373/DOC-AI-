"""
TICO Document Review Tool
Flask web application for AI-powered document review.
"""

import os
import uuid
import json
import threading
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import functools

# Load .env when present (dev convenience; prod uses real env vars).
try:
    from dotenv import load_dotenv as _ld
    _ld(os.path.join(os.path.dirname(__file__), ".env"))
except ImportError:
    pass

from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename
from supabase import create_client, Client

from doc_parser import parse_document, parse_excel, get_document_summary
from page_locator import enrich_pages
from review_engine import (
    create_ollama_client,
    create_failover_client,
    test_connection,
    review_document,
    REVIEW_CATEGORIES,
    SEVERITY_LEVELS,
)
from report_generator import generate_excel_report

app = Flask(__name__)
# Restrict CORS via env in production (comma-separated origins); defaults to open
# for local/dev so behavior is unchanged until ALLOWED_ORIGINS is set.
_allowed_origins = os.environ.get("ALLOWED_ORIGINS", "*")
CORS(app, origins="*" if _allowed_origins.strip() == "*" else [o.strip() for o in _allowed_origins.split(",") if o.strip()])

# Cap upload size (default 50 MB) to avoid memory-exhaustion uploads.
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "50")) * 1024 * 1024


def _safe_under(directory, filename):
    """Resolve `filename` under `directory`, refusing path traversal. None if unsafe."""
    base = os.path.realpath(directory)
    candidate = os.path.realpath(os.path.join(base, os.path.basename(filename)))
    if os.path.commonpath([candidate, base]) != base:
        return None
    return candidate

# Store active reviews and reports in memory
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
REPORTS_DIR = os.path.join(os.path.dirname(__file__), "reports")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

# Background execution: in-process threads (single-worker). Phase 2 replaces this
# with a durable Redis+RQ queue + a reviews DB table for crash-safe, multi-worker
# operation. The dead Huey/Redis scaffolding has been removed.

# Supabase Configuration
# Use Service Role key for backend (bypasses RLS, is authoritative).
# Falls back to Anon key so dev setups still work.
SUPABASE_URL = os.environ.get("VITE_SUPABASE_URL") or os.environ.get("SUPABASE_URL")
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
SUPABASE_KEY = SUPABASE_SERVICE_KEY or os.environ.get("VITE_SUPABASE_ANON_KEY") or os.environ.get("SUPABASE_ANON_KEY")
SUPABASE_JWT_SECRET = os.environ.get("SUPABASE_JWT_SECRET")

supabase: Client = None
if SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
        print("[*] Supabase: connected")
    except Exception as e:
        print(f"Supabase init error: {e}")
else:
    print("[*] Supabase: not configured — using local JSON state only")


# ── JWT Auth helper ─────────────────────────────────────────────────────────
def _get_user_id_from_token():
    """
    Extract user_id from the Supabase Bearer JWT in the request.
    Returns None when auth is not configured (on-prem/local mode) —
    caller treats None as "unauthenticated but allowed in local mode".
    """
    if not SUPABASE_URL:
        return None  # local/on-prem mode — no auth required

    auth_header = request.headers.get("Authorization", "")
    if not auth_header.startswith("Bearer "):
        return None
    token = auth_header[7:].strip()

    # Prefer local JWT decode when the secret is configured (no network round-trip).
    if SUPABASE_JWT_SECRET:
        try:
            import jwt as pyjwt
            payload = pyjwt.decode(token, SUPABASE_JWT_SECRET, algorithms=["HS256"],
                                   options={"verify_aud": False})
            return payload.get("sub")
        except Exception:
            return None

    # Fallback: ask Supabase to verify (adds a small round-trip).
    if supabase:
        try:
            user = supabase.auth.get_user(token)
            return user.user.id if user and user.user else None
        except Exception:
            return None
    return None


def require_auth(fn):
    """
    Decorator: gate an endpoint behind Supabase auth.
    When Supabase is not configured (SUPABASE_URL unset) the endpoint is
    allowed through so on-prem deployments work without login.
    """
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        if SUPABASE_URL:  # auth is enabled
            uid = _get_user_id_from_token()
            if not uid:
                return jsonify({"error": "Unauthorized"}), 401
        return fn(*args, **kwargs)
    return wrapper


# ── Per-review Supabase upsert ───────────────────────────────────────────────
def _supabase_upsert_review(review_id, user_id, data):
    """Write/update a single review row in the Supabase reviews table."""
    if not supabase:
        return
    try:
        row = {
            "id": review_id,
            "user_id": user_id or "00000000-0000-0000-0000-000000000000",
            **{k: data[k] for k in (
                "status", "progress", "message", "review_mode",
                "document_info", "findings", "summary", "report_filename",
                "pdf_filename", "engine_status", "error", "original_filepath",
            ) if k in data},
        }
        supabase.table("reviews").upsert(row).execute()
    except Exception as e:
        print(f"Supabase review upsert error: {e}")

# In-memory fallback (now with Supabase sync)
STATE_FILE = os.path.join(UPLOAD_DIR, "review_state.json")

def _load_store():
    store = {}
    # 1. Try to load from local file
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r') as f:
                store = json.load(f)
        except Exception:
            store = {}
    
    # 2. If nothing in local file OR local file missing, try Supabase (Cloud Persistence)
    if not store and supabase:
        try:
            response = supabase.table("review_state").select("*").execute()
            if response.data:
                for row in response.data:
                    store[row["id"]] = row["data"]
                # Save locally so subsequent reads are fast
                _save_store(store)
        except Exception as e:
            print(f"Supabase load error: {e}")
            
    return store

def _save_store(store):
    # Save to local file
    try:
        with open(STATE_FILE, 'w') as f:
            json.dump(store, f)
    except Exception:
        pass
    
    # Sync to Supabase if available
    if supabase:
        try:
            # We store the entire store as a JSON blob for simplicity in this migration
            # but ideally we'd use a row-per-review table.
            # Here we just iterate and upsert active reviews.
            for rid, data in store.items():
                if data.get("status") in ["done", "error"]:
                    # These might already be in history, but we ensure sync
                    pass
                supabase.table("review_state").upsert({
                    "id": rid,
                    "data": data,
                    "updated_at": "now()"
                }).execute()
        except Exception as e:
            print(f"Supabase sync error: {e}")


@app.route("/")
def index():
    return render_template("index.html")


def _requires_api_key(host):
    """Only the Ollama Cloud host needs an API key; local/self-hosted/on-prem don't."""
    return "ollama.com" in (host or "").lower()


@app.route("/api/check-ollama", methods=["POST"])
def check_ollama():
    """Test connection to an Ollama host (local or cloud)."""
    data = request.get_json()
    api_key = data.get("api_key", "")
    host = data.get("host", "https://ollama.com")

    if not api_key and _requires_api_key(host):
        return jsonify({"success": False, "error": "API key is required for cloud Ollama"})

    result = test_connection(api_key, host)
    return jsonify(result)


@app.route("/api/models", methods=["POST"])
def list_models():
    """List available models from an Ollama host (local or cloud)."""
    data = request.get_json()
    api_key = data.get("api_key", "")
    host = data.get("host", "https://ollama.com")

    if not api_key and _requires_api_key(host):
        return jsonify({"success": False, "error": "API key is required for cloud Ollama"})

    result = test_connection(api_key, host)
    return jsonify(result)


FEEDBACK_EMAIL = os.environ.get("FEEDBACK_EMAIL", "")

@app.route("/api/feedback", methods=["POST"])
def submit_feedback():
    """Accept user feedback with optional image, store in Supabase, and send email."""
    user_email = request.form.get("user_email", "unknown")
    feedback_type = request.form.get("type", "general")
    message = request.form.get("message", "")
    image_file = request.files.get("image")

    if not message.strip():
        return jsonify({"success": False, "error": "Feedback message is empty"})

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Upload image to Supabase Storage if provided
    image_url = None
    image_bytes = None
    image_filename = None
    if image_file and image_file.filename:
        image_bytes = image_file.read()
        image_filename = f"fb_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{image_file.filename}"
        if supabase:
            try:
                supabase.storage.from_("feedback-images").upload(
                    image_filename, image_bytes,
                    {"content-type": image_file.content_type or "image/png"}
                )
                image_url = supabase.storage.from_("feedback-images").get_public_url(image_filename)
                print(f"Feedback image uploaded: {image_url}")
            except Exception as e:
                print(f"Supabase image upload error: {e}")

    # Store in Supabase
    if supabase:
        try:
            supabase.table("feedback").insert({
                "user_email": user_email,
                "type": feedback_type,
                "message": message,
                "image_url": image_url,
                "created_at": timestamp,
            }).execute()
        except Exception as e:
            print(f"Supabase feedback insert error: {e}")

    # Send email notification
    try:
        smtp_host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
        smtp_port = int(os.environ.get("SMTP_PORT", "587"))
        smtp_user = os.environ.get("SMTP_USER", "")
        smtp_pass = os.environ.get("SMTP_PASS", "")

        if smtp_user and smtp_pass:
            from email.mime.base import MIMEBase
            from email import encoders

            msg = MIMEMultipart()
            msg["From"] = smtp_user
            msg["To"] = FEEDBACK_EMAIL
            msg["Subject"] = f"[DOC-AI Feedback] {feedback_type.upper()} from {user_email}"

            body = f"""New feedback received on DOC-AI Platform

From: {user_email}
Type: {feedback_type.upper()}
Time: {timestamp}
Screenshot: {image_url if image_url else "None"}

Message:
{message}
"""
            msg.attach(MIMEText(body, "plain"))

            # Attach image if provided
            if image_bytes and image_filename:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(image_bytes)
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={image_filename}")
                msg.attach(part)

            with smtplib.SMTP(smtp_host, smtp_port) as server:
                server.starttls()
                server.login(smtp_user, smtp_pass)
                server.sendmail(smtp_user, FEEDBACK_EMAIL, msg.as_string())
            print(f"Feedback email sent to {FEEDBACK_EMAIL}")
        else:
            print(f"SMTP not configured. Feedback stored in Supabase only.")
    except Exception as e:
        print(f"Email send error: {e}")

    return jsonify({"success": True, "message": "Feedback received"})


def _category_name_for_mode(category_key, review_mode=None):
    return REVIEW_CATEGORIES.get(category_key, {}).get("name", category_key)


def _category_icon_for_mode(category_key, review_mode=None):
    return REVIEW_CATEGORIES.get(category_key, {}).get("icon", "")


def _categories_for_mode(review_mode=None):
    return {
        k: {"name": v["name"], "icon": v["icon"], "description": v["description"]}
        for k, v in REVIEW_CATEGORIES.items()
    }


def _run_review_in_background(review_id, filepath, original_filename, api_key, host, model, review_mode="pro", file_type="doc", vision_model=None, user_id=None):
    """Background worker that runs the full document review."""
    store = _load_store()
    try:
        # Parse document
        if review_id not in store:
            store[review_id] = {}
        
        parsing_update = {"status": "parsing", "message": "Parsing document...", "progress": 10}
        store[review_id].update(parsing_update)
        _save_store(store)
        _supabase_upsert_review(review_id, user_id, parsing_update)

        if file_type == "excel":
            parsed = parse_excel(filepath)
        else:
            parsed = parse_document(filepath)
            # Replace heuristic page numbers with PDF-accurate ones when a
            # renderer (Gotenberg/LibreOffice) is available; safe no-op otherwise.
            try:
                enrich_pages(parsed, filepath, pdf_save_dir=REPORTS_DIR)
            except Exception as e:
                print(f"Page enrichment skipped: {e}")
        store = _load_store()
        store[review_id].update({
            "progress": 20,
            "message": f"Document parsed: {parsed['statistics']['total_words']} words, {parsed['statistics']['total_sections']} sections. Starting AI review..."
        })
        _save_store(store)

        # Create Ollama client (supports failover with multiple keys)
        api_keys = [k.strip() for k in api_key.split(",") if k.strip()]
        if len(api_keys) > 1:
            client = create_failover_client(api_keys, host)
            print(f"[Failover] Using {len(api_keys)} API keys with automatic rotation.")
        else:
            client = create_ollama_client(api_keys[0], host)

        # Progress callback
        def progress_cb(msg, explicit_pct=None):
            s = _load_store()
            if review_id not in s: return
            if explicit_pct is not None:
                new_progress = explicit_pct
            else:
                current = s[review_id].get("progress", 20)
                new_progress = min(current + 5, 88)
            s[review_id].update({
                "progress": new_progress,
                "message": msg,
                "status": "reviewing"
            })
            _save_store(s)

        # Run review — single unified engine. `review_mode` acts as a strictness
        # knob (normal < pro < max); max is pro + a high-confidence filter.
        engine_status = {}
        findings = review_document(
            client, model, parsed,
            progress_callback=progress_cb,
            review_mode=review_mode,
            vision_model=vision_model,
            status_out=engine_status,
        )

        # Generate report
        store = _load_store()
        store[review_id].update({
            "progress": 92,
            "message": f"Found {len(findings)} issues. Generating Excel report..."
        })
        _save_store(store)

        mode_suffix = {"normal": "Normal", "pro": "Pro", "max": "Max"}.get(review_mode, "Pro")
        report_base = secure_filename(os.path.splitext(original_filename)[0]) or "document"
        report_filename = (
            f"Review_Report_{report_base}"
            f"_{mode_suffix}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        report_path = os.path.join(REPORTS_DIR, report_filename)
        generate_excel_report(findings, original_filename, report_path)

        # Build stats
        severity_counts = {}
        category_counts = {}
        for f in findings:
            sev = f["severity"]
            severity_counts[sev] = severity_counts.get(sev, 0) + 1
            cat = f["category"]
            category_counts[cat] = category_counts.get(cat, 0) + 1

        # Mark done
        # Add default status and fix_type to each finding
        for f in findings:
            if "status" not in f:
                f["status"] = "OPEN"
            if "fix_type" not in f:
                f["fix_type"] = f.get("fix_type", "MANUAL")

        store = _load_store()
        store[review_id].update({
            "status": "done",
            "message": "Review complete!",
            "progress": 100,
            "report_filename": report_filename,
            "pdf_filename": parsed.get("metadata", {}).get("pdf_filename"),
            "review_mode": review_mode,
            "original_filepath": filepath,
            "engine_status": engine_status,
            "document_info": {
                "filename": original_filename,
                "words": parsed["statistics"]["total_words"],
                "sections": parsed["statistics"]["total_sections"],
                "tables": parsed["statistics"]["total_tables"],
                "images": parsed["statistics"]["total_images"],
            },
            "findings": findings,
            "summary": {
                "total_findings": len(findings),
                "severity_counts": severity_counts,
                "category_counts": {
                    k: {
                        "count": v,
                        "name": _category_name_for_mode(k, review_mode),
                        "icon": _category_icon_for_mode(k, review_mode),
                    }
                    for k, v in category_counts.items()
                },
            },
            "categories": _categories_for_mode(review_mode),
            "severity_levels": {
                k: {"label": v["label"], "color": v["color"]}
                for k, v in SEVERITY_LEVELS.items()
            },
        })
        _save_store(store)
        _supabase_upsert_review(review_id, user_id, store[review_id])

    except Exception as e:
        store = _load_store()
        if review_id in store:
            err_update = {"status": "error", "message": str(e), "progress": 0}
            store[review_id].update(err_update)
            _save_store(store)
            _supabase_upsert_review(review_id, user_id, err_update)
    finally:
        # Keep uploaded file for 24h for potential auto-fix application
        # Schedule cleanup after 24h instead of immediate deletion
        def _delayed_cleanup(path, delay_seconds=86400):
            import time
            time.sleep(delay_seconds)
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass
        cleanup_thread = threading.Thread(target=_delayed_cleanup, args=(filepath,), daemon=True)
        cleanup_thread.start()


@app.route("/api/review", methods=["POST"])
@require_auth
def start_review():
    """Upload a document and start the review process in background."""
    api_key = request.form.get("api_key", "")
    host = request.form.get("host", "https://ollama.com")
    model = request.form.get("model", "")
    vision_model = request.form.get("vision_model", "") or None
    review_mode = request.form.get("review_mode", "pro")
    file_type = request.form.get("file_type", "doc")

    if not api_key and _requires_api_key(host):
        return jsonify({"success": False, "error": "API key is required for cloud Ollama"})
    if not model:
        return jsonify({"success": False, "error": "Model selection is required"})
    if review_mode not in {"normal", "pro", "max"}:
        return jsonify({"success": False, "error": f"Invalid review mode: {review_mode}"})

    if "document" not in request.files:
        return jsonify({"success": False, "error": "No document file provided"})

    file = request.files["document"]
    if not file.filename:
        return jsonify({"success": False, "error": "No file selected"})

    if file_type == "excel":
        if not file.filename.lower().endswith((".xlsx", ".xls")):
            return jsonify({"success": False, "error": "Only .xlsx or .xls files are supported for Excel mode"})
    else:
        if not file.filename.lower().endswith((".docx", ".doc")):
            return jsonify({"success": False, "error": "Only .docx or .doc files are supported for Document mode"})

    # Save uploaded file under a sanitized name (prevents path traversal, B1).
    review_id = str(uuid.uuid4())[:8]
    user_id = _get_user_id_from_token()
    safe_name = secure_filename(file.filename) or "document"
    filename = f"{review_id}_{safe_name}"
    filepath = os.path.join(UPLOAD_DIR, filename)
    file.save(filepath)

    # Initialize progress
    initial = {"status": "starting", "message": "Uploading document...", "progress": 5}
    store = _load_store()
    store[review_id] = initial
    _save_store(store)
    _supabase_upsert_review(review_id, user_id, initial)

    # Dispatch to a background thread so progress polling works.
    thread = threading.Thread(
        target=_run_review_in_background,
        args=(review_id, filepath, file.filename, api_key, host, model, review_mode, file_type, vision_model, user_id),
        daemon=True,
    )
    thread.start()

    return jsonify({"success": True, "review_id": review_id})


@app.route("/api/progress/<review_id>")
@require_auth
def get_progress(review_id):
    """Get the progress/result of a review."""
    store = _load_store()
    if review_id not in store:
        return jsonify({"status": "unknown", "message": "Review not found", "progress": 0})

    data = store[review_id]

    # If done, return full results
    if data["status"] == "done":
        return jsonify(data)

    # If error, return error
    if data["status"] == "error":
        return jsonify(data)

    # Otherwise return progress
    return jsonify({
        "status": data["status"],
        "message": data["message"],
        "progress": data["progress"],
    })


@app.route("/api/update-finding/<review_id>", methods=["POST"])
@require_auth
def update_finding(review_id):
    """Update a finding's status (Accept/Reject/Working)."""
    data = request.get_json()
    finding_id = data.get("finding_id")
    new_status = data.get("status", "OPEN")

    if new_status not in ["OPEN", "WORKING", "CLOSED", "IGNORE", "N/A"]:
        return jsonify({"success": False, "error": f"Invalid status: {new_status}"}), 400

    store = _load_store()
    if review_id not in store or store[review_id].get("status") != "done":
        return jsonify({"success": False, "error": "Review not found or not complete"}), 404

    findings = store[review_id].get("findings", [])
    updated = False
    for f in findings:
        if f.get("id") == finding_id:
            f["status"] = new_status
            updated = True
            break

    if not updated:
        return jsonify({"success": False, "error": f"Finding {finding_id} not found"}), 404

    store[review_id]["findings"] = findings
    _save_store(store)
    return jsonify({"success": True, "finding_id": finding_id, "status": new_status})


@app.route("/api/apply-fixes/<review_id>", methods=["POST"])
@require_auth
def apply_fixes(review_id):
    """Apply auto-fixable findings to a copy of the document."""
    store = _load_store()
    if review_id not in store or store[review_id].get("status") != "done":
        return jsonify({"success": False, "error": "Review not found or not complete"}), 404

    data = request.get_json() or {}
    finding_ids = data.get("finding_ids")  # None = apply all auto-fixable

    findings = store[review_id].get("findings", [])
    original_filepath = store[review_id].get("original_filepath")

    if not original_filepath or not os.path.exists(original_filepath):
        return jsonify({"success": False, "error": "Original document no longer available. Please re-upload."}), 404

    try:
        from doc_fixer import apply_fixes as do_apply_fixes
        result = do_apply_fixes(original_filepath, findings, finding_ids)
        if result["success"]:
            fixed_filename = result["fixed_filename"]
            return jsonify({
                "success": True,
                "fixed_filename": fixed_filename,
                "changes_applied": result["changes_applied"],
                "changes_skipped": result["changes_skipped"],
            })
        else:
            return jsonify({"success": False, "error": result.get("error", "Unknown error")})
    except ImportError:
        return jsonify({"success": False, "error": "Auto-fix module not available yet."})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/download/<report_filename>")
@require_auth
def download_report(report_filename):
    """Download a generated Excel report (regenerated with latest statuses)."""
    # Find the review that owns this report
    store = _load_store()
    owner_review = None
    for rid, rdata in store.items():
        if rdata.get("report_filename") == report_filename:
            owner_review = rdata
            break

    report_path = _safe_under(REPORTS_DIR, report_filename)
    if not report_path:
        return jsonify({"error": "Invalid report name"}), 400

    # If we found the review, regenerate with latest statuses (all modes use one report).
    if owner_review and owner_review.get("findings"):
        try:
            findings = owner_review["findings"]
            doc_name = owner_review.get("document_info", {}).get("filename", "Unknown")
            generate_excel_report(findings, doc_name, report_path)
        except Exception as e:
            print(f"Report regeneration error: {e}")

    if os.path.exists(report_path):
        return send_file(
            report_path,
            as_attachment=True,
            download_name=os.path.basename(report_filename),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    return jsonify({"error": "Report not found"}), 404


@app.route("/api/download-fixed/<filename>")
@require_auth
def download_fixed_doc(filename):
    """Download the auto-fixed document."""
    fixed_path = _safe_under(REPORTS_DIR, filename)
    if not fixed_path:
        return jsonify({"error": "Invalid file name"}), 400
    if os.path.exists(fixed_path):
        return send_file(
            fixed_path,
            as_attachment=True,
            download_name=os.path.basename(filename),
        )
    return jsonify({"error": "Fixed document not found"}), 404


@app.route("/api/pdf/<review_id>")
@require_auth
def serve_pdf(review_id):
    """Serve the Gotenberg/soffice-rendered PDF for the PDF.js viewer."""
    store = _load_store()
    data = store.get(review_id, {})
    if not data:
        return jsonify({"error": "Review not found"}), 404

    pdf_filename = data.get("pdf_filename")
    if not pdf_filename:
        return jsonify({"error": "No rendered PDF available for this review — page locator used heuristic mode"}), 404

    pdf_path = _safe_under(REPORTS_DIR, pdf_filename)
    if not pdf_path or not os.path.exists(pdf_path):
        return jsonify({"error": "PDF file not found on disk"}), 404

    return send_file(pdf_path, mimetype="application/pdf")


if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  TICO Document Review Tool")
    print("  Open http://localhost:5000 in your browser")
    print("=" * 60 + "\n")
    app.run(debug=True, port=5000)
