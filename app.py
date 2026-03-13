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
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from supabase import create_client, Client

from doc_parser import parse_document, parse_excel, get_document_summary
from review_engine import (
    create_ollama_client,
    test_connection,
    review_document,
    REVIEW_CATEGORIES,
    SEVERITY_LEVELS,
)
from report_generator import generate_excel_report

app = Flask(__name__)
CORS(app)

# Store active reviews and reports in memory
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
REPORTS_DIR = os.path.join(os.path.dirname(__file__), "reports")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

# Supabase Configuration (Pull from Env)
SUPABASE_URL = os.environ.get("VITE_SUPABASE_URL")
SUPABASE_KEY = os.environ.get("VITE_SUPABASE_ANON_KEY") # Use Service Role if possible in production

supabase: Client = None
if SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        print(f"Supabase init error: {e}")

# In-memory fallback (now with Supabase sync)
STATE_FILE = os.path.join(UPLOAD_DIR, "review_state.json")

def _load_store():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

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


@app.route("/api/check-ollama", methods=["POST"])
def check_ollama():
    """Test connection to Ollama Cloud API."""
    data = request.get_json()
    api_key = data.get("api_key", "")
    host = data.get("host", "https://ollama.com")

    if not api_key:
        return jsonify({"success": False, "error": "API key is required"})

    result = test_connection(api_key, host)
    return jsonify(result)


@app.route("/api/models", methods=["POST"])
def list_models():
    """List available models from Ollama Cloud."""
    data = request.get_json()
    api_key = data.get("api_key", "")
    host = data.get("host", "https://ollama.com")

    if not api_key:
        return jsonify({"success": False, "error": "API key is required"})

    result = test_connection(api_key, host)
    return jsonify(result)


FEEDBACK_EMAIL = "yash.badgujar@getmysolutions.in"

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

    # Save image if provided
    image_path = None
    if image_file and image_file.filename:
        feedback_dir = os.path.join(os.path.dirname(__file__), "feedback_images")
        os.makedirs(feedback_dir, exist_ok=True)
        safe_name = f"fb_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{image_file.filename}"
        image_path = os.path.join(feedback_dir, safe_name)
        image_file.save(image_path)

    # Store in Supabase
    if supabase:
        try:
            supabase.table("feedback").insert({
                "user_email": user_email,
                "type": feedback_type,
                "message": message,
                "has_image": image_path is not None,
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
Has Screenshot: {"Yes" if image_path else "No"}

Message:
{message}
"""
            msg.attach(MIMEText(body, "plain"))

            # Attach image if provided
            if image_path and os.path.exists(image_path):
                with open(image_path, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(image_path)}")
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


def _run_review_in_background(review_id, filepath, original_filename, api_key, host, model, review_mode="pro", file_type="doc"):
    """Background worker that runs the full document review."""
    store = _load_store()
    try:
        # Parse document
        if review_id not in store:
            store[review_id] = {}
        
        store[review_id].update({
            "status": "parsing",
            "message": "Parsing document...",
            "progress": 10
        })
        _save_store(store)

        if file_type == "excel":
            parsed = parse_excel(filepath)
        else:
            parsed = parse_document(filepath)
        store = _load_store()
        store[review_id].update({
            "progress": 20,
            "message": f"Document parsed: {parsed['statistics']['total_words']} words, {parsed['statistics']['total_sections']} sections. Starting AI review..."
        })
        _save_store(store)

        # Create Ollama client
        client = create_ollama_client(api_key, host)

        # Progress callback
        def progress_cb(msg):
            s = _load_store()
            if review_id not in s: return
            current = s[review_id].get("progress", 20)
            new_progress = min(current + 5, 88)
            s[review_id].update({
                "progress": new_progress,
                "message": msg,
                "status": "reviewing"
            })
            _save_store(s)

        # Run review
        findings = review_document(client, model, parsed, progress_callback=progress_cb, review_mode=review_mode)

        # Generate report
        store = _load_store()
        store[review_id].update({
            "progress": 92,
            "message": f"Found {len(findings)} issues. Generating Excel report..."
        })
        _save_store(store)

        mode_suffix = "Normal" if review_mode == "normal" else "Pro"
        report_filename = (
            f"Review_Report_{os.path.splitext(original_filename)[0]}"
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
        store = _load_store()
        store[review_id].update({
            "status": "done",
            "message": "Review complete!",
            "progress": 100,
            "report_filename": report_filename,
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
                        "name": REVIEW_CATEGORIES.get(k, {}).get("name", k),
                        "icon": REVIEW_CATEGORIES.get(k, {}).get("icon", ""),
                    }
                    for k, v in category_counts.items()
                },
            },
            "categories": {
                k: {"name": v["name"], "icon": v["icon"], "description": v["description"]}
                for k, v in REVIEW_CATEGORIES.items()
            },
            "severity_levels": {
                k: {"label": v["label"], "color": v["color"]}
                for k, v in SEVERITY_LEVELS.items()
            },
        })
        _save_store(store)

    except Exception as e:
        store = _load_store()
        if review_id in store:
            store[review_id].update({
                "status": "error",
                "message": str(e),
                "progress": 0,
            })
            _save_store(store)
    finally:
        # Clean up uploaded file
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except Exception:
            pass


@app.route("/api/review", methods=["POST"])
def start_review():
    """Upload a document and start the review process in background."""
    api_key = request.form.get("api_key", "")
    host = request.form.get("host", "https://ollama.com")
    model = request.form.get("model", "")
    review_mode = request.form.get("review_mode", "pro")
    file_type = request.form.get("file_type", "doc")

    if not api_key:
        return jsonify({"success": False, "error": "API key is required"})
    if not model:
        return jsonify({"success": False, "error": "Model selection is required"})

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

    # Save uploaded file
    review_id = str(uuid.uuid4())[:8]
    filename = f"{review_id}_{file.filename}"
    filepath = os.path.join(UPLOAD_DIR, filename)
    file.save(filepath)

    # Initialize progress
    store = _load_store()
    store[review_id] = {
        "status": "starting",
        "message": "Uploading document...",
        "progress": 5,
    }
    _save_store(store)

    # Start background thread
    thread = threading.Thread(
        target=_run_review_in_background,
        args=(review_id, filepath, file.filename, api_key, host, model, review_mode, file_type),
        daemon=True,
    )
    thread.start()

    return jsonify({"success": True, "review_id": review_id})


@app.route("/api/progress/<review_id>")
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


@app.route("/api/download/<report_filename>")
def download_report(report_filename):
    """Download a generated Excel report."""
    report_path = os.path.join(REPORTS_DIR, report_filename)
    if os.path.exists(report_path):
        return send_file(
            report_path,
            as_attachment=True,
            download_name=report_filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    return jsonify({"error": "Report not found"}), 404


if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  TICO Document Review Tool")
    print("  Open http://localhost:5000 in your browser")
    print("=" * 60 + "\n")
    app.run(debug=True, port=5000)
