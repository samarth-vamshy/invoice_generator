"""End-to-end routing checks at the active Teams SDK handler boundary."""

import asyncio
import random
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, Mock

from microsoft_teams.api import AdaptiveCardInvokeActivity, MessageActivity
from microsoft_teams.cards import AdaptiveCard

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from domain import BatchResult, InvoiceRequest, StoredInvoice
from invoice_engine import generate_invoice
from invoice_routes import InvoiceTeamsRouter
from invoice_settings import Settings
from sharepoint_service import SearchRecord


class FakeContext:
    def __init__(self, activity):
        self.activity = activity
        self.sent = []

    async def send(self, message):
        self.sent.append(message)


def message(text=None, value=None):
    return FakeContext(MessageActivity.model_validate({
        "id": "turn-1", "from": {"id": "tester", "name": "Tester"},
        "recipient": {"id": "bot"},
        "conversation": {"id": "conversation-1", "conversationType": "personal"},
        "text": text, "value": value,
    }))


def action(data, verb=None):
    return FakeContext(AdaptiveCardInvokeActivity.model_validate({
        "id": "turn-2", "from": {"id": "tester", "name": "Tester"},
        "recipient": {"id": "bot"},
        "conversation": {"id": "conversation-1", "conversationType": "personal"},
        "value": {"action": {"type": "Action.Execute", "verb": verb or data.get("action"), "data": data}},
    }))


def card_json(card):
    assert isinstance(card, AdaptiveCard)
    return card.model_dump(by_alias=True, exclude_none=True)


class TeamsInvoiceFlowTests(unittest.TestCase):
    def setUp(self):
        self.fallback = AsyncMock()
        self.sharepoint = SimpleNamespace(library_url="https://contoso.sharepoint.com/InvoiceLibrary")
        self.sharepoint.search_invoice_metadata = Mock(return_value=[])
        self.sharepoint.list_batch_invoices = Mock(return_value=[])
        self.generation = SimpleNamespace(generate_custom=Mock(), generate_random=Mock())
        self.router = InvoiceTeamsRouter(
            model=None,
            fallback=self.fallback,
            settings_override=Settings.from_env({}),
            sharepoint=self.sharepoint,
            generation=self.generation,
        )

    def test_greeting_and_home_card_actions_use_teams_execute(self):
        greeting = message("Hi")
        asyncio.run(self.router.handle_message(greeting))
        self.assertEqual(greeting.sent[0], "Hi Tester, welcome to the Invoice Generator.\nWhat would you like to do?")
        home = card_json(greeting.sent[1])
        self.assertEqual([entry["type"] for entry in home["actions"]], ["Action.Execute"] * 3)
        self.assertEqual([entry["data"]["action"] for entry in home["actions"]], ["generate_invoices", "search_invoices", "help"])
        self.fallback.assert_not_awaited()

        generate = action({"action": "generate_invoices"})
        asyncio.run(self.router.handle_card_action(generate))
        self.assertEqual(card_json(generate.sent[0])["body"][0]["text"], "Generate Invoices")
        search = action({"action": "search_invoices"})
        asyncio.run(self.router.handle_card_action(search))
        self.assertEqual(search.sent[0], "Ask me anything about your invoices.")
        help_ctx = action({"action": "help"})
        asyncio.run(self.router.handle_card_action(help_ctx))
        self.assertIn("SharePoint Invoice Library", help_ctx.sent[0])
        self.fallback.assert_not_awaited()

    def test_review_then_confirm_uses_entered_details_once(self):
        values = {
            "action": "review_invoice_request", "date_from": "2026-09-01", "date_to": "2026-09-30",
            "min_amount": "1000", "max_amount": "2000", "seller_name": "Seller Ltd",
            "seller_address": "Seller Road", "seller_tax_id": "SELLER-TAX",
            "buyer_name": "Buyer Ltd", "buyer_address": "Buyer Street", "buyer_tax_id": "BUYER-TAX",
            "tax_type": "IGST", "tax_rate": "18",
        }
        review = action(values)
        asyncio.run(self.router.handle_card_action(review))
        self.generation.generate_custom.assert_not_called()
        self.assertIn("Seller name: Seller Ltd", review.sent[0])
        self.assertIn("Buyer tax ID: BUYER-TAX", review.sent[0])
        self.assertEqual(card_json(review.sent[1])["actions"][0]["data"]["action"], "generate_confirmed_invoices")

        request = InvoiceRequest.from_form({**values, "company_mode": "Custom"})
        invoice = generate_invoice(request, "INV_STANDARD_TESTING_20260929_ABC_001", "BATCH-20260929-ABC", "Tester", random.Random(7))
        stored = StoredInvoice(invoice, "item-1", "https://contoso.sharepoint.com/invoice.pdf")
        self.generation.generate_custom.return_value = BatchResult(invoice.batch_id, 1, 1, 1, (), (stored,))
        token = self.router.pending[("conversation-1", "tester")]["token"]
        confirm = action({"action": "generate_confirmed_invoices", "token": token})
        asyncio.run(self.router.handle_card_action(confirm))
        self.generation.generate_custom.assert_called_once()
        self.assertEqual(card_json(confirm.sent[0])["actions"][0]["url"], stored.web_url)
        self.assertEqual(self.generation.generate_custom.call_args.args[0].seller_name, "Seller Ltd")
        replay = action({"action": "generate_confirmed_invoices", "token": token})
        asyncio.run(self.router.handle_card_action(replay))
        self.generation.generate_custom.assert_called_once()
        self.assertIn("no longer active", replay.sent[0])

    def test_invalid_review_returns_feedback_without_generation(self):
        review = action({"action": "review_invoice_request", "date_from": "2026-09-01", "date_to": "2026-09-30", "min_amount": "bad", "max_amount": "2000", "seller_name": "Seller", "buyer_name": "Buyer"})
        asyncio.run(self.router.handle_card_action(review))
        self.assertIn("Minimum amount", review.sent[0])
        self.assertEqual(card_json(review.sent[1])["body"][0]["text"], "Generate an invoice")
        self.generation.generate_custom.assert_not_called()

    def test_random_card_action_generates_batch_and_invalid_count_returns_form(self):
        invalid = action({"action": "generate_random_batch", "invoice_count": "0", "purpose": "Testing"})
        asyncio.run(self.router.handle_card_action(invalid))
        self.assertIn("Number of invoices", invalid.sent[0])
        self.assertEqual(card_json(invalid.sent[1])["body"][0]["text"], "Generate Random Invoices")
        self.generation.generate_random.assert_not_called()

        self.generation.generate_random.return_value = BatchResult("BATCH-20260929-ABC", 1, 0, 0, (), ())
        valid = action({"action": "generate_random_batch", "invoice_count": 1, "purpose": "Testing"})
        asyncio.run(self.router.handle_card_action(valid))
        self.generation.generate_random.assert_called_once()
        request = self.generation.generate_random.call_args.args[0]
        self.assertEqual((request.count, request.purpose), (1, "Testing"))
        self.assertIn("storage failures", card_json(valid.sent[0])["body"][0]["text"])

    def test_edit_and_cancel_do_not_generate(self):
        values = {
            "action": "review_invoice_request", "date_from": "2026-09-01", "date_to": "2026-09-30",
            "min_amount": "1000", "max_amount": "2000", "seller_name": "Seller Ltd",
            "buyer_name": "Buyer Ltd", "tax_type": "IGST", "tax_rate": "18",
        }
        asyncio.run(self.router.handle_card_action(action(values)))
        token = self.router.pending[("conversation-1", "tester")]["token"]
        edit = action({"action": "edit", "token": token})
        asyncio.run(self.router.handle_card_action(edit))
        form = card_json(edit.sent[0])
        self.assertEqual(form["body"][0]["text"], "Generate an invoice")
        seller = next(field for field in form["body"] if field.get("id") == "seller_name")
        self.assertEqual(seller["value"], "Seller Ltd")
        self.assertNotIn(("conversation-1", "tester"), self.router.pending)
        cancel = action({"action": "cancel"})
        asyncio.run(self.router.handle_card_action(cancel))
        self.assertEqual(cancel.sent[0], "Invoice creation cancelled.")
        self.assertEqual(card_json(cancel.sent[-1])["body"][0]["text"], "Invoice Generator")
        self.generation.generate_custom.assert_not_called()

    def test_search_uses_sharepoint_and_unrelated_text_uses_llm(self):
        record = SearchRecord("item-1", "https://contoso.sharepoint.com/invoice.pdf", {"Invoice Number": "INV_GST_UAT_001", "Purpose": "UAT", "Total Amount": 1000})
        self.sharepoint.search_invoice_metadata.return_value = [record]
        search = message("Show me UAT invoices")
        asyncio.run(self.router.handle_message(search))
        self.sharepoint.search_invoice_metadata.assert_called_once_with("Show me UAT invoices", 10)
        self.assertEqual(card_json(search.sent[0])["body"][2]["actions"][0]["url"], record.web_url)
        self.fallback.assert_not_awaited()
        prompt = action({"action": "search_invoices"})
        asyncio.run(self.router.handle_card_action(prompt))
        followup = message("What invoices did I generate?")
        asyncio.run(self.router.handle_message(followup))
        self.sharepoint.search_invoice_metadata.assert_called_with("What invoices did I generate?", 10)
        self.fallback.assert_not_awaited()
        general = message("What is a proforma invoice?")
        asyncio.run(self.router.handle_message(general))
        self.fallback.assert_awaited_once()


if __name__ == "__main__":
    unittest.main()
