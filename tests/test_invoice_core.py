"""Business logic tests that do not need Graph, Microsoft 365, or Azure OpenAI."""

import json
import random
import sys
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from domain import InvoiceRequest
from errors import InvoiceValidationError
from invoice_engine import calculate_totals, generate_invoice, validate_invoice
from local_invoice_service import LocalInvoiceService
from domain import LineItem


def request(**overrides):
    form = {
        "invoice_type": "GST", "purpose": "UAT", "date_from": "2026-09-01", "date_to": "2026-09-30",
        "min_amount": "20000", "max_amount": "75000", "currency": "INR", "tax_type": "CGST/SGST",
        "tax_rate": "18", "line_item_mode": "Mixed", "max_line_items": "5", "company_mode": "Generated",
    }
    form.update(overrides)
    return InvoiceRequest.from_form(form)


class InvoiceCoreTests(unittest.TestCase):
    def test_gst_split_uses_decimal_and_reconciles(self):
        items = (LineItem("Service", 1, Decimal("1000.05"), Decimal("1000.05")),)
        totals = calculate_totals(items, "CGST/SGST", Decimal("18"))
        self.assertEqual(totals.tax.cgst, Decimal("90.00"))
        self.assertEqual(totals.tax.sgst, Decimal("90.00"))
        self.assertEqual(totals.grand_total, Decimal("1180.05"))

    def test_generated_invoice_meets_range_and_reconciles(self):
        spec = request()
        for seed in range(30):
            invoice = generate_invoice(spec, f"INV_GST_UAT_20260928_{seed:03d}", "BATCH-20260928-001", "tester", random.Random(seed))
            validate_invoice(invoice, spec)
            self.assertGreaterEqual(invoice.totals.grand_total, spec.min_amount)
            self.assertLessEqual(invoice.totals.grand_total, spec.max_amount)
            self.assertEqual(sum(item.amount for item in invoice.line_items), invoice.totals.subtotal)

    def test_fixed_total_can_be_met(self):
        spec = request(min_amount="20000", max_amount="20000", tax_type="IGST")
        invoice = generate_invoice(spec, "INV_GST_UAT_20260928_001", "BATCH-20260928-001", "tester", random.Random(7))
        self.assertEqual(invoice.totals.grand_total, Decimal("20000.00"))

    def test_invalid_card_values_are_rejected(self):
        for values in (
            {"min_amount": "NaN"}, {"min_amount": "-1"}, {"date_from": "2026-10-01"},
            {"tax_type": "None"}, {"tax_rate": "0"}, {"count": "51"}, {"max_line_items": "0"},
            {"generation_mode": "Single", "count": "3"}, {"output_destination": "Local disk"},
        ):
            with self.subTest(values=values), self.assertRaises(InvoiceValidationError):
                request(**values)

    def test_custom_batch_count_is_validated(self):
        spec = request(generation_mode="Batch", count="5")
        self.assertEqual((spec.generation_mode, spec.count, spec.output_destination), ("Batch", 5, "PDF in Invoice Library"))

    def test_entered_company_fields_override_generated_mode(self):
        spec = request(
            seller_name="Entered Seller", seller_address="Seller Address", seller_tax_id="SELLER-123",
            buyer_name="Entered Buyer", buyer_address="Buyer Address", buyer_tax_id="BUYER-456",
        )
        self.assertEqual(spec.company_mode, "Custom")
        invoice = generate_invoice(spec, "INV_GST_UAT_20260928_001", "BATCH-20260928-001", "tester", random.Random(7))
        self.assertEqual((invoice.seller.name, invoice.seller.address, invoice.seller.tax_id), ("Entered Seller", "Seller Address", "SELLER-123"))
        self.assertEqual((invoice.buyer.name, invoice.buyer.address, invoice.buyer.tax_id), ("Entered Buyer", "Buyer Address", "BUYER-456"))
        self.assertEqual(invoice.totals.tax.rate, Decimal("18.00"))

    def test_custom_batch_rejects_amount_range_without_distinct_totals(self):
        with self.assertRaisesRegex(InvoiceValidationError, "Widen the amount range"):
            request(
                generation_mode="Batch", count="2", min_amount="1000", max_amount="1000",
                seller_name="Entered Seller", buyer_name="Entered Buyer",
            )

    def test_local_result_keeps_pdf_and_structured_data_with_unique_names(self):
        with tempfile.TemporaryDirectory() as folder:
            service = LocalInvoiceService(Path(folder))
            spec = request(min_amount="1000", max_amount="2000")
            first = service.generate_one(spec, "tester")
            second = service.generate_one(spec, "tester")
            self.assertNotEqual(first.invoice.invoice_number, second.invoice.invoice_number)
            self.assertTrue(first.pdf_path.read_bytes().startswith(b"%PDF"))
            self.assertEqual(json.loads(first.json_path.read_text(encoding="utf-8"))["totals"]["grand_total"], str(first.invoice.totals.grand_total))

    def test_existing_pdf_survives_a_name_collision(self):
        with tempfile.TemporaryDirectory() as folder:
            service = LocalInvoiceService(Path(folder))
            first = service.generate_one(request(), "tester")
            original = first.pdf_path.read_bytes()
            service._next_number = lambda _request: (first.invoice.invoice_number, first.invoice.batch_id)
            from errors import PdfGenerationError
            with self.assertRaises(PdfGenerationError):
                service.generate_one(request(), "tester")
            self.assertEqual(first.pdf_path.read_bytes(), original)


if __name__ == "__main__":
    unittest.main()
