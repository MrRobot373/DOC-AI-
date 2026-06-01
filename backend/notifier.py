"""
Email notifications — sent when a review completes (or fails).

Reuses the same SMTP env the feedback endpoint uses (SMTP_HOST/PORT/USER/PASS).
Safe no-op when SMTP isn't configured. Never raises into the caller.
"""

from __future__ import annotations

import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def _smtp_configured():
    return bool(os.environ.get("SMTP_USER") and os.environ.get("SMTP_PASS"))


def _send(to_email, subject, body):
    if not to_email or not _smtp_configured():
        return False
    try:
        host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
        port = int(os.environ.get("SMTP_PORT", "587"))
        user = os.environ["SMTP_USER"]
        pwd = os.environ["SMTP_PASS"]

        msg = MIMEMultipart()
        msg["From"] = user
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        with smtplib.SMTP(host, port) as server:
            server.starttls()
            server.login(user, pwd)
            server.sendmail(user, to_email, msg.as_string())
        print(f"[notify] sent review email to {to_email}")
        return True
    except Exception as e:
        print(f"[notify] email send failed: {e}")
        return False


def send_review_complete_email(to_email, doc_name, summary, report_link=None):
    """summary: {total_findings, severity_counts:{CRITICAL,MAJOR,MINOR}}"""
    counts = (summary or {}).get("severity_counts", {})
    total = (summary or {}).get("total_findings", 0)
    crit = counts.get("CRITICAL", 0)
    major = counts.get("MAJOR", 0)
    minor = counts.get("MINOR", 0)
    body = f"""Your DOC-AI review is ready.

Document: {doc_name}
Findings: {total}  (Critical: {crit}, Major: {major}, Minor: {minor})
"""
    if report_link:
        body += f"\nOpen the report: {report_link}\n"
    subject = f"[DOC-AI] Review complete — {doc_name} ({crit} critical, {total} total)"
    return _send(to_email, subject, body)


def send_review_error_email(to_email, doc_name, error):
    subject = f"[DOC-AI] Review failed — {doc_name}"
    body = f"Your DOC-AI review of '{doc_name}' could not be completed.\n\nError: {error}\n"
    return _send(to_email, subject, body)
