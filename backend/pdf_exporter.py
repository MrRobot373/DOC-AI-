"""
PDF export of a findings report (paginated, print-friendly).

Mirrors the Excel report content but as a PDF for sharing/printing. Uses reportlab.
"""

from __future__ import annotations

from datetime import datetime
import html

SEV_COLOR = {"CRITICAL": "#c0392b", "MAJOR": "#e67e22", "MINOR": "#b7950b"}
SEV_ORDER = {"CRITICAL": 0, "MAJOR": 1, "MINOR": 2}


def findings_to_pdf(findings, doc_name, out_path):
    """Render findings to a paginated PDF at out_path. Returns out_path."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    )

    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    h2 = styles["Heading2"]
    body = ParagraphStyle("body", parent=styles["Normal"], fontSize=8, leading=10)

    doc = SimpleDocTemplate(out_path, pagesize=A4,
                            leftMargin=14 * mm, rightMargin=14 * mm,
                            topMargin=14 * mm, bottomMargin=14 * mm)
    elems = []
    elems.append(Paragraph(f"DOC-AI Review Report", title_style))
    elems.append(Paragraph(html.escape(f"{doc_name}"), h2))
    elems.append(Paragraph(f"Generated: {datetime.now():%Y-%m-%d %H:%M}", body))
    elems.append(Spacer(1, 6 * mm))

    # Severity summary
    counts = {"CRITICAL": 0, "MAJOR": 0, "MINOR": 0}
    for f in findings:
        counts[f.get("severity", "MINOR")] = counts.get(f.get("severity", "MINOR"), 0) + 1
    elems.append(Paragraph(
        f"<b>{len(findings)} findings</b> — Critical: {counts['CRITICAL']}, "
        f"Major: {counts['MAJOR']}, Minor: {counts['MINOR']}", body))
    elems.append(Spacer(1, 4 * mm))

    # Findings table
    header = ["#", "Sev", "Page", "Category", "Comment", "Fix"]
    rows = [header]
    ordered = sorted(findings, key=lambda f: SEV_ORDER.get(f.get("severity", "MINOR"), 3))
    for i, f in enumerate(ordered, 1):
        rows.append([
            str(i),
            f.get("severity", "")[:4],
            str(f.get("page", "-")),
            Paragraph(html.escape(str(f.get("category", "")).replace("_", " ").title()), body),
            Paragraph(html.escape((f.get("comment", "") or "")[:600]), body),
            Paragraph(html.escape((f.get("fix", "") or "")[:300]), body),
        ])

    table = Table(rows, colWidths=[8 * mm, 12 * mm, 12 * mm, 30 * mm, 75 * mm, 45 * mm], repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a1a2e")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#cccccc")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5f7")]),
    ]
    for ri, f in enumerate(ordered, 1):
        c = SEV_COLOR.get(f.get("severity", "MINOR"))
        if c:
            style.append(("TEXTCOLOR", (1, ri), (1, ri), colors.HexColor(c)))
    table.setStyle(TableStyle(style))
    elems.append(table)

    doc.build(elems)
    return out_path
