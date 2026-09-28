"""ReportLab presentation of a validated invoice."""

from __future__ import annotations

from html import escape
from io import BytesIO

from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from domain import Invoice
from errors import PdfGenerationError


NAVY = colors.HexColor("#102A43")
BLUE = colors.HexColor("#1769AA")
PALE = colors.HexColor("#EAF2F8")
MUTED = colors.HexColor("#52606D")


def _money(invoice: Invoice, amount: object) -> str:
    return f"{invoice.currency} {amount:,.2f}"


def _footer(canvas, document) -> None:
    canvas.saveState()
    width, _ = document.pagesize
    canvas.setStrokeColor(colors.HexColor("#CBD2D9"))
    canvas.line(18 * mm, 19 * mm, width - 18 * mm, 19 * mm)
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(MUTED)
    canvas.drawString(18 * mm, 14 * mm, "Generated for testing and demonstration. Not valid for payment.")
    canvas.drawRightString(width - 18 * mm, 14 * mm, f"Page {document.page}")
    canvas.restoreState()


def render_pdf(invoice: Invoice) -> bytes:
    """Render an already validated invoice; no accounting occurs here."""
    try:
        buffer = BytesIO()
        document = SimpleDocTemplate(buffer, pagesize=(210 * mm, 297 * mm), rightMargin=18 * mm, leftMargin=18 * mm, topMargin=18 * mm, bottomMargin=28 * mm)
        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name="InvoiceTitle", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=22, leading=26, textColor=NAVY, spaceAfter=7 * mm))
        styles.add(ParagraphStyle(name="Muted", parent=styles["Normal"], textColor=MUTED, fontSize=9, leading=13))
        styles.add(ParagraphStyle(name="Right", parent=styles["Normal"], alignment=TA_RIGHT, fontSize=9, leading=13))
        styles.add(ParagraphStyle(name="Cell", parent=styles["Normal"], fontSize=9, leading=12))
        p = lambda value, style="Normal": Paragraph(escape(str(value)).replace("\n", "<br/>"), styles[style])
        story = [p("INVOICE", "InvoiceTitle"), p("TEST / DEMO DOCUMENT - NOT VALID FOR PAYMENT", "Muted"), Spacer(1, 6 * mm)]
        header = Table([
            [p("Invoice number", "Muted"), p(invoice.invoice_number, "Right")],
            [p("Invoice date", "Muted"), p(invoice.invoice_date.isoformat(), "Right")],
            [p("Due date", "Muted"), p(invoice.due_date.isoformat(), "Right")],
            [p("Purpose", "Muted"), p(invoice.purpose, "Right")],
        ], colWidths=[75 * mm, 99 * mm])
        header.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), PALE), ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#D7E2EB")), ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
        story += [header, Spacer(1, 9 * mm)]
        seller = f"{invoice.seller.name}\n{invoice.seller.address}" + (f"\nTax ID: {invoice.seller.tax_id}" if invoice.seller.tax_id else "")
        buyer = f"{invoice.buyer.name}\n{invoice.buyer.address}" + (f"\nTax ID: {invoice.buyer.tax_id}" if invoice.buyer.tax_id else "")
        parties = Table([[p("FROM", "Muted"), p("BILL TO", "Muted")], [p(seller), p(buyer)]], colWidths=[87 * mm, 87 * mm])
        parties.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, 0), 5), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
        story += [parties, Spacer(1, 9 * mm)]
        rate = invoice.totals.tax.rate.normalize()
        tax_label = "-" if invoice.totals.tax.tax_type == "None" else f"{rate}%"
        rows = [[p("DESCRIPTION", "Muted"), p("QTY", "Muted"), p("UNIT PRICE", "Muted"), p("TAX", "Muted"), p("AMOUNT", "Muted")]]
        rows += [[p(item.description, "Cell"), str(item.quantity), _money(invoice, item.unit_price), tax_label, _money(invoice, item.amount)] for item in invoice.line_items]
        item_table = Table(rows, colWidths=[62 * mm, 13 * mm, 31 * mm, 18 * mm, 50 * mm], repeatRows=1)
        item_table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), PALE), ("LINEBELOW", (0, 0), (-1, 0), 1, BLUE), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7), ("ALIGN", (1, 1), (-1, -1), "RIGHT"), ("ALIGN", (1, 1), (1, -1), "CENTER"), ("ALIGN", (3, 1), (3, -1), "CENTER")]))
        story += [item_table, Spacer(1, 8 * mm)]
        totals = invoice.totals
        summary = [["Subtotal", _money(invoice, totals.subtotal)]]
        if totals.tax.cgst:
            summary.append([f"CGST ({(totals.tax.rate / 2).normalize()}%)", _money(invoice, totals.tax.cgst)])
            summary.append([f"SGST ({(totals.tax.rate / 2).normalize()}%)", _money(invoice, totals.tax.sgst)])
        if totals.tax.igst:
            summary.append([f"IGST ({totals.tax.rate.normalize()}%)", _money(invoice, totals.tax.igst)])
        summary.append(["GRAND TOTAL", _money(invoice, totals.grand_total)])
        summary_table = Table(summary, colWidths=[74 * mm, 55 * mm], hAlign="RIGHT")
        summary_table.setStyle(TableStyle([("ALIGN", (1, 0), (1, -1), "RIGHT"), ("LINEABOVE", (0, -1), (-1, -1), 1, NAVY), ("TEXTCOLOR", (0, -1), (-1, -1), NAVY), ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("TOPPADDING", (0, -1), (-1, -1), 9)]))
        story += [KeepTogether([summary_table, Spacer(1, 10 * mm), p(f"Batch: {invoice.batch_id} | Generated for {invoice.purpose.lower()} use only.", "Muted")])]
        document.build(story, onFirstPage=_footer, onLaterPages=_footer)
        return buffer.getvalue()
    except Exception as exc:
        raise PdfGenerationError("The invoice PDF could not be rendered.") from exc
