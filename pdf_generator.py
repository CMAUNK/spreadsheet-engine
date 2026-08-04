"""Geração do relatório PDF da operação SmartSheet."""
from __future__ import annotations

from datetime import date
from io import BytesIO
from typing import Any

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


def generate_pdf(title: str, sheet_name: str, rows: list[dict[str, Any]], generated_date: date) -> bytes:
    buffer = BytesIO()
    document = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=1.8 * cm, rightMargin=1.8 * cm,
                                 topMargin=1.8 * cm, bottomMargin=1.8 * cm)
    styles = getSampleStyleSheet()
    table_data = [["Referência", "Valor"]] + [[str(row.get("Referência", "")), str(row.get("Valor", ""))] for row in rows]
    table = Table(table_data, colWidths=[11 * cm, 5 * cm], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A5F")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#B8C4D1")),
        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F7FA")]),
        ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    elements = [Paragraph(title, styles["Title"]), Spacer(1, 0.3 * cm),
                Paragraph(f"Aba: {sheet_name} &nbsp;&nbsp;|&nbsp;&nbsp; Data: {generated_date.strftime('%d/%m/%Y')}", styles["Normal"]),
                Spacer(1, 0.5 * cm), table]
    document.build(elements)
    return buffer.getvalue()
