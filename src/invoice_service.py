"""One validated invoice/PDF/SharePoint pipeline for custom and random batches."""

from __future__ import annotations

import hashlib
import json
import logging
import random
import re
from datetime import date, timedelta
from decimal import Decimal
from uuid import uuid4

from config import Settings
from domain import BatchResult, FailedInvoice, INVOICE_TYPES, Invoice, InvoiceRequest, LINE_ITEM_MODES, RandomGenerationRequest
from errors import InvoiceGenerationError, InvoiceValidationError, PdfGenerationError, SharePointError
from invoice_engine import generate_invoice, validate_invoice
from pdf_renderer import render_pdf
from sharepoint_service import GenerationActor, SharePointInvoiceService


logger = logging.getLogger(__name__)


def _code(value: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "_", value.upper()).strip("_")


def _payload_signature(invoice: Invoice) -> str:
    payload = {
        "type": invoice.invoice_type,
        "purpose": invoice.purpose,
        "date": invoice.invoice_date.isoformat(),
        "seller": invoice.seller.name,
        "buyer": invoice.buyer.name,
        "items": [(item.description, item.quantity, str(item.unit_price)) for item in invoice.line_items],
        "tax_type": invoice.totals.tax.tax_type,
        "tax_rate": str(invoice.totals.tax.rate),
        "total": str(invoice.totals.grand_total),
    }
    return hashlib.sha256(json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")).hexdigest()


class InvoiceGenerationService:
    def __init__(self, settings: Settings, sharepoint: SharePointInvoiceService):
        self.settings = settings
        self.sharepoint = sharepoint

    def _random_invoice_request(self, random_request: RandomGenerationRequest, rng: random.Random) -> InvoiceRequest:
        invoice_type = rng.choice(INVOICE_TYPES)
        tax_type = rng.choice(("IGST", "CGST/SGST"))
        tax_rate = rng.choice(("5", "12", "18"))
        today = date.today()
        return InvoiceRequest.from_form({
            "invoice_type": invoice_type,
            "purpose": random_request.purpose,
            "date_from": (today - timedelta(days=30)).isoformat(),
            "date_to": today.isoformat(),
            "min_amount": str(self.settings.random_min_amount),
            "max_amount": str(self.settings.random_max_amount),
            "currency": self.settings.default_currency,
            "tax_type": tax_type,
            "tax_rate": tax_rate,
            "line_item_mode": rng.choice(LINE_ITEM_MODES),
            "max_line_items": str(min(5, self.settings.max_line_items)),
            "company_mode": "Generated",
            "count": "1",
        }, self.settings.max_batch_size, self.settings.max_line_items, self.settings.max_amount)

    def generate_custom(self, request: InvoiceRequest, actor: GenerationActor) -> BatchResult:
        normalized = InvoiceRequest.from_form(request.to_form(), self.settings.max_batch_size, self.settings.max_line_items, self.settings.max_amount)
        if normalized.company_mode != "Custom":
            raise InvoiceValidationError("Customized invoices need seller and buyer names. Use Generate Random for generated companies.")
        available_cent_amounts = int((normalized.max_amount - normalized.min_amount) * 100) + 1
        if normalized.count > available_cent_amounts:
            raise InvoiceValidationError("Widen the amount range so each invoice in the batch can have a different total.")
        return self._run_batch(normalized.count, actor, lambda _rng: normalized, random_mode=False)

    def generate_random(self, request: RandomGenerationRequest, actor: GenerationActor) -> BatchResult:
        normalized = RandomGenerationRequest.from_form(
            {"invoice_count": str(request.count), "purpose": request.purpose},
            self.settings.max_batch_size,
            self.settings.default_purpose,
        )
        return self._run_batch(normalized.count, actor, lambda rng: self._random_invoice_request(normalized, rng), random_mode=True)

    def _run_batch(self, count: int, actor: GenerationActor, request_factory, random_mode: bool) -> BatchResult:
        if count < 1 or count > self.settings.max_batch_size:
            raise InvoiceValidationError(f"Invoice count must be between 1 and {self.settings.max_batch_size}.")
        person_lookup_id = self.sharepoint.preflight(actor)
        stamp = date.today().strftime("%Y%m%d")
        batch_suffix = uuid4().hex[:12].upper()
        batch_id = f"BATCH-{stamp}-{batch_suffix}"
        rng = random.Random()
        seen_invoice_numbers: set[str] = set()
        seen_party_pairs: set[tuple[str, str]] = set()
        seen_payload_hashes: set[str] = set()
        seen_amounts: set[Decimal] = set()
        seen_tax_values: set[Decimal] = set()
        stored = []
        failures = []
        generated_count = 0
        fixed_date: date | None = None
        logger.info("batch_started batch_id=%s requested_count=%s random=%s", batch_id, count, random_mode)
        for index in range(1, count + 1):
            stage = "generation"
            try:
                for retry in range(self.settings.random_generation_max_retries):
                    spec = request_factory(rng)
                    invoice_number = f"INV_{_code(spec.invoice_type)}_{_code(spec.purpose)}_{stamp}_{batch_suffix}_{index:03d}"
                    if invoice_number in seen_invoice_numbers:
                        raise InvoiceGenerationError("A duplicate invoice number was generated.")
                    invoice = generate_invoice(spec, invoice_number, batch_id, actor.display_name, rng, fixed_details=not random_mode, fixed_date=fixed_date)
                    validate_invoice(invoice, spec)
                    signature = _payload_signature(invoice)
                    pair = (invoice.seller.name.casefold(), invoice.buyer.name.casefold())
                    tax_value = invoice.totals.tax.cgst + invoice.totals.tax.sgst + invoice.totals.tax.igst
                    unique = (
                        pair not in seen_party_pairs
                        and signature not in seen_payload_hashes
                        and invoice.totals.grand_total not in seen_amounts
                        and tax_value not in seen_tax_values
                    ) if random_mode else invoice.totals.grand_total not in seen_amounts
                    if unique:
                        break
                    logger.info("invoice_generation_retry batch_id=%s item=%s mode=%s retry_count=%s", batch_id, index, "random" if random_mode else "custom", retry + 1)
                else:
                    raise InvoiceGenerationError("Could not create a unique invoice amount within the configured retry limit. Widen the amount range and try again.")
                if not random_mode and fixed_date is None:
                    fixed_date = invoice.invoice_date
                seen_invoice_numbers.add(invoice_number)
                seen_party_pairs.add(pair)
                seen_payload_hashes.add(signature)
                seen_amounts.add(invoice.totals.grand_total)
                seen_tax_values.add(tax_value)
                stage = "pdf"
                pdf_bytes = render_pdf(invoice)
                generated_count += 1
                logger.info("pdf_generated batch_id=%s invoice_number=%s bytes=%s", batch_id, invoice_number, len(pdf_bytes))
                stage = "sharepoint"
                stored.append(self.sharepoint.upload_invoice(invoice, pdf_bytes, person_lookup_id))
                logger.info("invoice_stored batch_id=%s invoice_number=%s", batch_id, invoice_number)
            except (InvoiceGenerationError, InvoiceValidationError, PdfGenerationError, SharePointError) as exc:
                logger.exception("invoice_batch_item_failed batch_id=%s item=%s stage=%s", batch_id, index, stage)
                failures.append(FailedInvoice(index=index, stage=stage, message=str(exc)))
            except Exception:
                logger.exception("invoice_batch_item_failed_unexpected batch_id=%s item=%s stage=%s", batch_id, index, stage)
                failures.append(FailedInvoice(index=index, stage=stage, message="An unexpected error occurred."))
        result = BatchResult(
            batch_id=batch_id,
            requested_count=count,
            generated_count=generated_count,
            uploaded_count=len(stored),
            failed_items=tuple(failures),
            sharepoint_records=tuple(stored),
        )
        logger.info(
            "batch_complete batch_id=%s requested=%s generated=%s uploaded=%s failed=%s",
            batch_id, count, result.generated_count, result.uploaded_count, result.failed_count,
        )
        return result
