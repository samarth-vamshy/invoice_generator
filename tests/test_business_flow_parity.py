"""GitHub business flow exercised through Teams activities and a fake Graph boundary."""

import asyncio
import json
import sys
import unittest
from datetime import datetime, timedelta, timezone
from decimal import Decimal
from pathlib import Path
from unittest.mock import AsyncMock, Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from microsoft_teams.api import ConversationUpdateActivity, MessageActivity
from cards import invoice_form_card, random_generation_card
from invoice_routes import InvoiceTeamsRouter
from invoice_service import InvoiceGenerationService
from sharepoint_service import SharePointInvoiceService
from test_invoice_teams_flow import FakeContext, action, card_json, message
from test_sharepoint_pipeline import FakeGraph, settings


class BusinessFlowParityTests(unittest.TestCase):
    def setUp(self):
        self.settings = settings()
        self.graph = FakeGraph()
        self.sharepoint = SharePointInvoiceService(self.settings, self.graph)
        self.generation = InvoiceGenerationService(self.settings, self.sharepoint)
        self.generated = []
        original_custom = self.generation.generate_custom

        def capture_custom(request, actor):
            result = original_custom(request, actor)
            self.generated.append(result)
            return result

        self.generation.generate_custom = Mock(side_effect=capture_custom)
        self.fallback = AsyncMock()
        self.router = InvoiceTeamsRouter(None, self.fallback, self.settings, self.sharepoint, self.generation)

    def click(self, card, title, values=None):
        button = next(item for item in card["actions"] if item["title"] == title)
        ctx = action({**button["data"], **(values or {})}, button["verb"])
        ctx.activity.from_.properties = {"email": "local@example.test"}
        asyncio.run(self.router.handle_card_action(ctx))
        self.fallback.assert_not_awaited()
        return ctx

    def custom_form(self):
        home = message("hello")
        asyncio.run(self.router.handle_message(home))
        methods = self.click(card_json(home.sent[-1]), "Generate Invoices")
        form = self.click(card_json(methods.sent[-1]), "Customize")
        return card_json(form.sent[-1])

    def values(self, form):
        values = {field["id"]: field.get("value", "") for field in form["body"] if "id" in field}
        values.update({
            "invoice_type": "GST", "purpose": "Demo", "generation_mode": "Batch", "count": "3",
            "date_from": "2026-09-01", "date_to": "2026-09-30", "min_amount": "1,000", "max_amount": "9,000",
            "seller_name": "Exact Seller", "seller_address": "Exact Seller Address", "seller_tax_id": "SELLER-123",
            "buyer_name": "Exact Buyer", "buyer_address": "Exact Buyer Address", "buyer_tax_id": "BUYER-456",
            "tax_type": "CGST/SGST", "tax_rate": "12", "currency": "USD", "line_item_mode": "Services",
        })
        return values

    def test_custom_cards_to_pdf_metadata_and_search_keep_all_entered_values(self):
        form = self.custom_form()
        values = self.values(form)
        review = self.click(form, "Review", values)
        self.assertFalse(self.graph.files, "Review must not render or upload a PDF")
        self.generation.generate_custom.assert_not_called()
        facts = {fact["title"]: fact["value"] for fact in card_json(review.sent[-1])["body"][2]["facts"]}
        self.assertEqual(facts["Seller address"], values["seller_address"])
        self.assertEqual(facts["Buyer tax ID"], values["buyer_tax_id"])
        self.assertEqual(facts["Tax rate"], "12%")
        result = self.click(card_json(review.sent[-1]), "Generate")
        batch = self.generated[-1]
        self.assertEqual((batch.requested_count, batch.generated_count, batch.uploaded_count), (3, 3, 3))
        invoices = [record.invoice for record in batch.sharepoint_records]
        for invoice in invoices:
            self.assertEqual((invoice.seller.name, invoice.seller.address, invoice.seller.tax_id), (values["seller_name"], values["seller_address"], values["seller_tax_id"]))
            self.assertEqual((invoice.buyer.name, invoice.buyer.address, invoice.buyer.tax_id), (values["buyer_name"], values["buyer_address"], values["buyer_tax_id"]))
            self.assertEqual((invoice.currency, invoice.totals.tax.tax_type, invoice.totals.tax.rate), ("USD", "CGST/SGST", Decimal("12")))
            self.assertLessEqual(Decimal("1000"), invoice.totals.grand_total)
            self.assertLessEqual(invoice.totals.grand_total, Decimal("9000"))
        self.assertEqual(len({invoice.invoice_date for invoice in invoices}), 1)
        self.assertEqual(len({tuple(item.description for item in invoice.line_items) for invoice in invoices}), 1)
        self.assertEqual(len({invoice.totals.grand_total for invoice in invoices}), 3)
        self.assertTrue(all(pdf.startswith(b"%PDF") for _, pdf in self.graph.files.values()))
        self.assertEqual({item["fields"]["Seller_Name"] for item in self.graph.items.values()}, {"Exact Seller"})
        batch_view = self.click(card_json(result.sent[-1]), "View Batch")
        self.assertEqual(sum(item["type"] == "ActionSet" for item in card_json(batch_view.sent[-1])["body"]), 3)
        for query, count in (("Show me Demo invoices", 3), ("Show invoices generated today", 3), (invoices[0].invoice_number, 1)):
            with self.subTest(query=query):
                ctx = message(query)
                asyncio.run(self.router.handle_message(ctx))
                self.assertEqual(sum(item["type"] == "ActionSet" for item in card_json(ctx.sent[-1])["body"]), count)
        self.fallback.assert_not_awaited()

    def test_random_partial_storage_result_and_view_batch_use_real_pipeline(self):
        self.graph.fail_metadata_update = 2
        ctx = action({"action": "generate_random"})
        asyncio.run(self.router.handle_card_action(ctx))
        result = self.click(card_json(ctx.sent[-1]), "Generate", {"invoice_count": 3, "purpose": "UAT"})
        card = card_json(result.sent[-1])
        facts = {fact["title"]: fact["value"] for fact in card["body"][1]["facts"]}
        self.assertEqual((facts["Generated"], facts["Stored"], facts["Failures"]), ("3", "2", "1"))
        self.assertEqual(len(self.graph.files), 2, "A PDF with failed metadata must be removed")
        batch = self.click(card, "View Batch")
        self.assertEqual(sum(item["type"] == "ActionSet" for item in card_json(batch.sent[-1])["body"]), 2)

    def test_generated_by_email_sources_survive_the_teams_transport(self):
        email = "identity@example.test"
        cases = (
            ({"email": email}, {}),
            ({"properties": {"email": email}}, {}),
            ({"properties": {"userPrincipalName": email}}, {}),
            ({}, {"userPrincipalName": email}),
            ({}, {"userEmail": email}),
        )
        identity_settings = settings(SHAREPOINT_USER_LOOKUP_IDS=json.dumps({email: 29}))
        service = SharePointInvoiceService(identity_settings, self.graph)
        for account_fields, channel_fields in cases:
            with self.subTest(account=account_fields, channel=channel_fields):
                ctx = FakeContext(MessageActivity.model_validate({
                    "id": "identity", "recipient": {"id": "bot"}, "conversation": {"id": "conversation-1"},
                    "from": {"id": "tester", "name": "Tester", **account_fields}, "channelData": channel_fields,
                }))
                actor = self.router._actor(ctx)
                self.assertEqual(actor.email, email)
                self.assertEqual(service.preflight(actor), 29)

    def test_cancel_and_home_can_leave_an_incomplete_form(self):
        for form in (invoice_form_card(), random_generation_card(50, "Testing")):
            for button in form["actions"]:
                if button["title"] in {"Cancel", "Home"}:
                    self.assertEqual(button["associatedInputs"], "none")
                    ctx = self.click(form, button["title"])
                    self.assertEqual(card_json(ctx.sent[-1])["body"][0]["text"], "Invoice Generator")
                else:
                    self.assertEqual(button["associatedInputs"], "auto")
        self.assertFalse(self.graph.files)

    def test_natural_language_hints_reach_custom_form_in_teams(self):
        ctx = message("<at>Invoice Bot</at> Generate 3 GST invoices for September 2026 UAT between INR 1000 and INR 9000")
        asyncio.run(self.router.handle_message(ctx))
        form = self.click(card_json(ctx.sent[-1]), "Customize")
        fields = {field["id"]: field.get("value") for field in card_json(form.sent[-1])["body"] if "id" in field}
        for key, value in {"count": "3", "generation_mode": "Batch", "purpose": "UAT", "invoice_type": "GST", "tax_rate": "18", "date_from": "2026-09-01", "date_to": "2026-09-30", "min_amount": "1000", "max_amount": "9000"}.items():
            self.assertEqual(fields[key], value, key)
        self.assertFalse(self.graph.files)

    def test_personal_conversation_start_attaches_home(self):
        ctx = FakeContext(ConversationUpdateActivity.model_validate({
            "id": "start", "from": {"id": "tester", "name": "Tester"}, "recipient": {"id": "bot"},
            "conversation": {"id": "conversation-1", "conversationType": "personal"}, "membersAdded": [{"id": "bot"}],
        }))
        asyncio.run(self.router.handle_conversation_update(ctx))
        self.assertIn("welcome to the Invoice Generator", ctx.sent[0])
        self.assertEqual(len(card_json(ctx.sent[-1])["actions"]), 3)

    def test_legacy_submit_messages_still_use_business_routes(self):
        for route, title in (("generate_invoices", "Generate Invoices"), ("generate_random", "Generate Random Invoices"), ("customize_invoice", "Generate an invoice"), ("go_home", "Invoice Generator")):
            with self.subTest(route=route):
                ctx = message("", value={"route": route})
                asyncio.run(self.router.handle_message(ctx))
                self.assertEqual(card_json(ctx.sent[-1])["body"][0]["text"], title)
        self.fallback.assert_not_awaited()
        self.assertFalse(self.graph.files)

    def test_review_cannot_be_confirmed_by_another_user_or_after_expiry(self):
        form = self.custom_form()
        review = self.click(form, "Review", self.values(form))
        button = card_json(review.sent[-1])["actions"][0]
        other_user = action(button["data"])
        other_user.activity.from_.id = "another-user"
        asyncio.run(self.router.handle_card_action(other_user))
        self.assertIn("no longer active", other_user.sent[0])
        self.router.pending[("conversation-1", "tester")]["created_at"] = datetime.now(timezone.utc) - timedelta(minutes=16)
        expired = self.click(card_json(review.sent[-1]), "Generate")
        self.assertIn("no longer active", expired.sent[0])
        self.generation.generate_custom.assert_not_called()
        self.assertFalse(self.graph.files)

    def test_missing_company_names_cannot_reach_generation(self):
        form = self.custom_form()
        values = self.values(form)
        values.update(seller_name="", buyer_name="")
        review = self.click(form, "Review", values)
        self.assertIn("both seller and buyer names", review.sent[0])
        self.assertFalse(self.router.pending)
        self.assertFalse(self.graph.files)

    def test_review_logs_values_and_failures_include_traceback(self):
        form = self.custom_form()
        with self.assertLogs("invoice_routes", level="INFO") as logged:
            self.click(form, "Review", self.values(form))
        log = "\n".join(logged.output)
        self.assertIn("selected_route=review_invoice_request", log)
        self.assertIn("'tax_rate': '12'", log)
        self.assertNotIn("Exact Seller Address", log)
        self.generation.generate_random = Mock(side_effect=RuntimeError("pipeline diagnostic"))
        ctx = action({"action": "generate_random_batch", "invoice_count": 1})
        with self.assertLogs("invoice_routes", level="ERROR") as error_log:
            asyncio.run(self.router.handle_card_action(ctx))
        self.assertTrue(any(record.exc_info for record in error_log.records))
        self.assertIn("Something went wrong", ctx.sent[0])
        self.assertFalse(self.graph.files)


if __name__ == "__main__":
    unittest.main()
