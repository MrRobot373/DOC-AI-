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

# In-memory store for review progress and results
review_store = {}


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


def _run_review_in_background(review_id, filepath, original_filename, api_key, host, model):
    """Background worker that runs the full document review."""
    try:
        # Parse document
        review_store[review_id]["status"] = "parsing"
        review_store[review_id]["message"] = "Parsing document..."
        review_store[review_id]["progress"] = 10

        parsed = parse_document(filepath)
        review_store[review_id]["progress"] = 20
        review_store[review_id]["message"] = (
            f"Document parsed: {parsed['statistics']['total_words']} words, "
            f"{parsed['statistics']['total_sections']} sections. Starting AI review..."
        )

        # Create Ollama client
        client = create_ollama_client(api_key, host)

        # Progress callback
        def progress_cb(msg):
            current = review_store[review_id]["progress"]
            # Increment progress gradually from 20 to 88
            new_progress = min(current + 5, 88)
            review_store[review_id]["progress"] = new_progress
            review_store[review_id]["message"] = msg
            review_store[review_id]["status"] = "reviewing"

        # Run review
        findings = review_document(client, model, parsed, progress_callback=progress_cb)

        # Generate report
        review_store[review_id]["progress"] = 92
        review_store[review_id]["message"] = f"Found {len(findings)} issues. Generating Excel report..."

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
        review_store[review_id].update({
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

    except Exception as e:
        review_store[review_id].update({
            "status": "error",
            "message": str(e),
            "progress": 0,
        })
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
    review_store[review_id] = {
        "status": "starting",
        "message": "Uploading document...",
        "progress": 5,
    }

    # Start background thread
    thread = threading.Thread(
        target=_run_review_in_background,
        args=(review_id, filepath, file.filename, api_key, host, model),
        daemon=True,
    )
    thread.start()

    return jsonify({"success": True, "review_id": review_id})


@app.route("/api/progress/<review_id>")
def get_progress(review_id):
    """Get the progress/result of a review."""
    if review_id not in review_store:
        return jsonify({"status": "unknown", "message": "Review not found", "progress": 0})

    data = review_store[review_id]

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
