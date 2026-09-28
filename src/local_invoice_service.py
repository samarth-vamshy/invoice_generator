"""Local milestone storage, kept separate from generation and rendering."""

from __future__ import annotations

import json
import logging
import random
import re
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from domain import Invoice, InvoiceRequest
from errors import PdfGenerationError
from invoice_engine import generate_invoice
from pdf_renderer import render_pdf


logger = logging.getLogger(__name__)
TYPE_CODES = {"Standard": "STD", "GST": "GST", "Service": "SVC", "Product": "PRD", "Proforma": "PRO", "Commercial": "COM"}
PURPOSE_CODES = {"Testing": "TEST", "UAT": "UAT", "Demo": "DEMO", "Training": "TRAIN", "Development": "DEV", "Other": "OTHER"}


@dataclass(frozen=True)
class LocalGenerationResult:
    invoice: Invoice
    pdf_path: Path
    json_path: Path


class LocalInvoiceService:
    def __init__(self, output_dir: Path):
        self.output_dir = output_dir
        self._lock = threading.Lock()

    def _next_number(self, request: InvoiceRequest) -> tuple[str, str]:
        stamp = datetime.now().strftime("%Y%m%d")
        prefix = f"INV_{TYPE_CODES[request.invoice_type]}_{PURPOSE_CODES[request.purpose]}_{stamp}_"
        matches = [int(match.group(1)) for path in self.output_dir.glob(f"{prefix}*.pdf") if (match := re.fullmatch(re.escape(prefix) + r"(\d+)\.pdf", path.name))]
        sequence = max(matches, default=0) + 1
        return f"{prefix}{sequence:03d}", f"BATCH-{stamp}-{sequence:03d}"

    def generate_one(self, request: InvoiceRequest, generated_by: str) -> LocalGenerationResult:
        """Generate one PDF and its structured JSON sidecar atomically enough for local use."""
        if request.count != 1:
            raise ValueError("This local milestone creates one invoice per request.")
        self.output_dir.mkdir(parents=True, exist_ok=True)
        with self._lock:
            invoice_number, batch_id = self._next_number(request)
            invoice = generate_invoice(request, invoice_number, batch_id, generated_by, random.Random())
            pdf_bytes = render_pdf(invoice)
            pdf_path = self.output_dir / f"{invoice_number}.pdf"
            json_path = self.output_dir / f"{invoice_number}.json"
            created_pdf = False
            created_json = False
            try:
                with pdf_path.open("xb") as file:
                    created_pdf = True
                    file.write(pdf_bytes)
                with json_path.open("x", encoding="utf-8") as file:
                    created_json = True
                    json.dump(invoice.to_dict(), file, indent=2)
            except (OSError, TypeError, ValueError) as exc:
                if created_pdf:
                    pdf_path.unlink(missing_ok=True)
                if created_json:
                    json_path.unlink(missing_ok=True)
                raise PdfGenerationError("The invoice could not be saved locally.") from exc
            logger.info("invoice_saved invoice_number=%s batch_id=%s", invoice_number, batch_id)
            return LocalGenerationResult(invoice, pdf_path, json_path)
