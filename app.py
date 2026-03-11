"""
TICO Document Review Tool
Flask web application for AI-powered document review.
"""

import os
import uuid
import json
import threading
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS

from doc_parser import parse_document, get_document_summary
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

# In-memory store (now with file persistence)
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
    try:
        with open(STATE_FILE, 'w') as f:
            json.dump(store, f)
    except Exception:
        pass


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


def _run_review_in_background(review_id, filepath, original_filename, api_key, host, model, review_mode="pro"):
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

        report_filename = (
            f"Review_Report_{os.path.splitext(original_filename)[0]}"
            f"_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
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

    if not api_key:
        return jsonify({"success": False, "error": "API key is required"})
    if not model:
        return jsonify({"success": False, "error": "Model selection is required"})

    if "document" not in request.files:
        return jsonify({"success": False, "error": "No document file provided"})

    file = request.files["document"]
    if not file.filename:
        return jsonify({"success": False, "error": "No file selected"})

    if not file.filename.lower().endswith((".docx", ".doc")):
        return jsonify({"success": False, "error": "Only .docx files are supported"})

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
        args=(review_id, filepath, file.filename, api_key, host, model, review_mode),
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
