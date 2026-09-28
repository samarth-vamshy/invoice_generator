"""Review and confirmation test using a fake SDK context."""

import asyncio
import random
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import agent
from cards import invoice_form_card
from domain import BatchResult, FailedInvoice, InvoiceRequest, StoredInvoice
from invoice_engine import generate_invoice
from sharepoint_service import SearchRecord
from routing import requested_invoice_count
from microsoft_agents.hosting.core import MemoryStorage, TurnState


class FakeState:
    def __init__(self):
        self.values = {}

    def get_value(self, name, default):
        return self.values.get(name, default())

    def set_value(self, name, value):
        self.values[name] = value

    def delete_value(self, name):
        self.values.pop(name, None)


class FakeContext:
    def __init__(self, value=None, text=None):
        self.activity = SimpleNamespace(type="message", name=None, text=text, value=value, id="turn-1", channel_id="msteams", conversation=SimpleNamespace(id="test-conversation"), from_property=SimpleNamespace(id="tester", name="Tester"))
        self.turn_state = {}
        self.sent = []

    async def send_activity(self, activity):
        self.sent.append(activity)


class AgentFlowTests(unittest.TestCase):
    def test_form_review_action_has_no_silent_regex_gate(self):
        form = invoice_form_card()
        self.assertEqual(form["actions"][0]["type"], "Action.Submit")
        self.assertEqual(form["actions"][0]["data"], {"route": "review_invoice_request", "company_mode": "Custom"})
        self.assertTrue(next(field for field in form["body"] if field.get("id") == "seller_name")["isRequired"])
        self.assertTrue(next(field for field in form["body"] if field.get("id") == "buyer_name")["isRequired"])
        self.assertFalse(any("regex" in field for field in form["body"] if field["type"].startswith("Input.")))

    def test_hi_uses_active_home_handler_and_attaches_card(self):
        context = FakeContext(text="Hi")
        with patch.object(agent, "answer_general_question", new_callable=AsyncMock) as llm:
            asyncio.run(agent.on_message(context, FakeState()))
            llm.assert_not_awaited()
        self.assertEqual(context.sent[0], "Hi Tester, welcome to the Invoice Generator.\nWhat would you like to do?")
        attachment = context.sent[1].attachments[0]
        self.assertEqual(attachment.content_type, "application/vnd.microsoft.card.adaptive")
        self.assertEqual([action["title"] for action in attachment.content["actions"]], ["Generate Invoices", "Search Invoices", "Help"])
        self.assertEqual([action["data"]["route"] for action in attachment.content["actions"]], ["generate_invoices", "search_invoices", "help"])

    def test_home_buttons_use_explicit_handlers(self):
        async def scenario():
            storage = MemoryStorage()
            state = TurnState.with_storage(storage)
            responses = []
            for action in ("generate_invoices", "search_invoices", "help"):
                context = FakeContext(value={"route": action})
                await state.load(context, storage)
                await agent.on_message(context, state)
                responses.append(context.sent)
            return responses

        with patch.object(agent, "answer_general_question", new_callable=AsyncMock) as llm:
            responses = asyncio.run(scenario())
            llm.assert_not_awaited()
        method_card = responses[0][0].attachments[0].content
        self.assertEqual(method_card["body"][0]["text"], "Generate Invoices")
        self.assertEqual([action["data"]["route"] for action in method_card["actions"]], ["generate_random", "customize_invoice", "go_home"])
        self.assertEqual(responses[1][0], "Ask me anything about your invoices.")
        self.assertEqual(responses[1][1].attachments[0].content["actions"][0]["data"]["route"], "go_home")
        self.assertIn("SharePoint Invoice Library", responses[2][0])

    def test_general_question_uses_llm_fallback(self):
        context = FakeContext(text="What is a proforma invoice?")
        with patch.object(agent, "answer_general_question", new_callable=AsyncMock, return_value="A preliminary invoice.") as llm:
            asyncio.run(agent.on_message(context, FakeState()))
            llm.assert_awaited_once()
        self.assertEqual(context.sent, ["A preliminary invoice."])

    def test_batch_text_is_recognized(self):
        self.assertEqual(requested_invoice_count("Generate 10 GST invoices for September"), 10)

    def test_review_then_confirm_and_replay_rejected(self):
        form = {"route": "review_invoice_request", "company_mode": "Generated", "date_from": "2026-09-01", "date_to": "2026-09-30", "min_amount": "1,000", "max_amount": "2,000", "seller_name": "Seller Ltd", "seller_address": "Seller Road", "seller_tax_id": "SELLER-TAX", "buyer_name": "Buyer Ltd", "buyer_address": "Buyer Street", "buyer_tax_id": "BUYER-TAX", "tax_type": "IGST", "tax_rate": "18"}
        request = InvoiceRequest.from_form(form)
        invoice = generate_invoice(request, "INV_STANDARD_TESTING_20260928_ABC_001", "BATCH-20260928-ABC", "Tester", random.Random(7))
        record = StoredInvoice(invoice, "item-1", "https://contoso.sharepoint.com/invoice.pdf")
        batch = BatchResult(invoice.batch_id, 1, 1, 1, (), (record,))

        async def scenario():
            state = FakeState()
            review_context = FakeContext(form)
            await agent.on_message(review_context, state)
            generator.generate_custom.assert_not_called()
            self.assertIn("Review your invoice:\nInvoice type: Standard", review_context.sent[0])
            self.assertIn("Seller name: Seller Ltd", review_context.sent[0])
            self.assertIn("Buyer tax ID: BUYER-TAX", review_context.sent[0])
            self.assertIn("Tax rate: 18%", review_context.sent[0])
            self.assertEqual(state.values[agent.PENDING_KEY]["form"]["company_mode"], "Custom")
            self.assertIn("No document has been created", review_context.sent[0])
            facts = review_context.sent[1].attachments[0].content["body"][2]["facts"]
            self.assertIn("Line item mode", [fact["title"] for fact in facts])
            self.assertEqual(review_context.sent[1].attachments[0].content["actions"][0]["data"]["route"], "generate_confirmed_invoices")
            token = state.values[agent.PENDING_KEY]["token"]
            confirm_context = FakeContext({"route": "generate_confirmed_invoices", "token": token})
            await agent.on_message(confirm_context, state)
            self.assertEqual(generator.generate_custom.call_count, 1)
            self.assertEqual(confirm_context.sent[0].attachments[0].content["actions"][0]["url"], record.web_url)
            self.assertNotIn(agent.PENDING_KEY, state.values)
            replay_context = FakeContext({"route": "generate_confirmed_invoices", "token": token})
            await agent.on_message(replay_context, state)
            self.assertEqual(generator.generate_custom.call_count, 1)

        with patch.object(agent, "generation_service") as generator, patch.object(agent, "sharepoint_service", SimpleNamespace(library_url="https://contoso.sharepoint.com/InvoiceLibrary")):
            generator.generate_custom.return_value = batch
            asyncio.run(scenario())

    def test_generation_method_routes_and_home(self):
        async def scenario():
            state = FakeState()
            random_context = FakeContext({"route": "generate_random"})
            await agent.on_message(random_context, state)
            self.assertEqual(random_context.sent[0].attachments[0].content["actions"][0]["data"]["route"], "generate_random_batch")
            customize_context = FakeContext({"route": "customize_invoice"})
            await agent.on_message(customize_context, state)
            self.assertEqual(customize_context.sent[0].attachments[0].content["body"][0]["text"], "Generate an invoice")
            self.assertEqual(customize_context.sent[0].attachments[0].content["actions"][-1]["data"]["route"], "go_home")
            home_context = FakeContext({"route": "go_home"})
            await agent.on_message(home_context, state)
            self.assertEqual(home_context.sent[1].attachments[0].content["body"][0]["text"], "Invoice Generator")
        asyncio.run(scenario())

    def test_random_generate_action_reports_partial_sharepoint_result(self):
        request = InvoiceRequest.from_form({"date_from": "2026-09-01", "date_to": "2026-09-30", "min_amount": "1000", "max_amount": "2000"})
        invoice = generate_invoice(request, "INV_STANDARD_TESTING_20260928_ABC_001", "BATCH-20260928-ABC", "Tester", random.Random(7))
        record = StoredInvoice(invoice, "item-1", "https://contoso.sharepoint.com/invoice.pdf")
        batch = BatchResult(invoice.batch_id, 3, 3, 1, (FailedInvoice(2, "sharepoint", "Metadata failed"), FailedInvoice(3, "sharepoint", "Upload failed")), (record,))
        context = FakeContext({"route": "generate_random_batch", "invoice_count": 3, "purpose": "Testing"})
        with patch.object(agent, "generation_service") as generator, patch.object(agent, "sharepoint_service", SimpleNamespace(library_url="https://contoso.sharepoint.com/InvoiceLibrary")), patch.object(agent, "answer_general_question", new_callable=AsyncMock) as llm:
            generator.generate_random.return_value = batch
            asyncio.run(agent.on_message(context, FakeState()))
            generator.generate_random.assert_called_once()
            llm.assert_not_awaited()
        card = context.sent[0].attachments[0].content
        self.assertIn("storage failures", card["body"][0]["text"])
        self.assertEqual([(fact["title"], fact["value"]) for fact in card["body"][1]["facts"]][-2:], [("Stored", "1"), ("Failures", "2")])
        self.assertEqual(card["actions"][0]["data"]["route"], "view_batch")

    def test_invalid_random_count_is_visible_and_does_not_generate(self):
        context = FakeContext({"route": "generate_random_batch", "invoice_count": "0", "purpose": "Testing"})
        with patch.object(agent, "generation_service") as generator:
            asyncio.run(agent.on_message(context, FakeState()))
            generator.generate_random.assert_not_called()
        self.assertIn("between 1 and", context.sent[0])
        self.assertEqual(context.sent[1].attachments[0].content["actions"][0]["data"]["route"], "generate_random_batch")

    def test_search_query_uses_sharepoint_metadata(self):
        record = SearchRecord("item-1", "https://contoso.sharepoint.com/invoice.pdf", {"Invoice Number": "INV_GST_UAT_20260928_001", "Purpose": "UAT", "Total Amount": 1000})
        context = FakeContext(text="Show me UAT invoices")
        with patch.object(agent, "sharepoint_service") as service, patch.object(agent, "answer_general_question", new_callable=AsyncMock) as llm:
            service.search_invoice_metadata.return_value = [record]
            asyncio.run(agent.on_message(context, FakeState()))
            service.search_invoice_metadata.assert_called_once()
            llm.assert_not_awaited()
        self.assertEqual(context.sent[0].attachments[0].content["body"][2]["actions"][0]["url"], record.web_url)

    def test_invalid_review_returns_visible_feedback(self):
        context = FakeContext({"route": "review_invoice_request", "date_from": "2026-09-01", "date_to": "2026-09-30", "min_amount": "bad", "max_amount": "2000", "seller_name": "Seller Ltd", "buyer_name": "Buyer Ltd"})
        state = FakeState()
        asyncio.run(agent.on_message(context, state))
        self.assertIn("Minimum amount", context.sent[0])
        self.assertEqual(context.sent[1].attachments[0].content["body"][0]["text"], "Generate an invoice")
        self.assertNotIn(agent.PENDING_KEY, state.values)

    def test_customize_requires_company_names_before_review(self):
        context = FakeContext({"route": "review_invoice_request", "date_from": "2026-09-01", "date_to": "2026-09-30", "min_amount": "1000", "max_amount": "2000"})
        state = FakeState()
        asyncio.run(agent.on_message(context, state))
        self.assertIn("both seller and buyer names", context.sent[0])
        self.assertEqual(context.sent[1].attachments[0].content["actions"][0]["data"]["company_mode"], "Custom")
        self.assertNotIn(agent.PENDING_KEY, state.values)


if __name__ == "__main__":
    unittest.main()
