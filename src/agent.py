"""Microsoft 365 Agents SDK transport handlers for Invoice Generator."""

from __future__ import annotations

import asyncio
import logging
import re
from datetime import datetime, timedelta, timezone
from uuid import uuid4

from microsoft_agents.activity import Activity, ActivityTypes, Attachment
from microsoft_agents.authentication.msal import MsalConnectionManager
from microsoft_agents.hosting.aiohttp import CloudAdapter
from microsoft_agents.hosting.core import AgentApplication, MemoryStorage, TurnContext, TurnState

from ai_client import answer_general_question
from cards import (
    generation_method_card, home_card, invoice_form_card, random_generation_card,
    result_card, review_card, review_summary_text, search_prompt_card, search_results_card,
)
from config import agents_sdk_config, settings
from domain import InvoiceRequest, RandomGenerationRequest
from errors import InvoiceGenerationError, InvoiceValidationError, SharePointError
from invoice_service import InvoiceGenerationService
from routing import prefill_from_text, requested_invoice_count, route_message
from sharepoint_service import GenerationActor, SharePointInvoiceService


logger = logging.getLogger(__name__)
logging.basicConfig(level=getattr(logging, settings.log_level), format="%(asctime)s %(levelname)s %(name)s %(message)s")

storage = MemoryStorage()
connection_manager = MsalConnectionManager(**agents_sdk_config)
adapter = CloudAdapter(connection_manager=connection_manager)
agent_app = AgentApplication[TurnState](storage=storage, adapter=adapter, connection_manager=connection_manager, **agents_sdk_config)
sharepoint_service = SharePointInvoiceService(settings)
generation_service = InvoiceGenerationService(settings, sharepoint_service)
PENDING_KEY = "UserState.pending_invoice"
PREFILL_KEY = "UserState.invoice_prefill"
HOME_ACTIONS = {"generate_invoices", "search_invoices", "help"}
REVIEW_ACTIONS = {"review", "review_invoice_request"}
CONFIRM_ACTIONS = {"confirm", "generate_confirmed_invoices"}
HELP_TEXT = "Choose Generate Invoices for random or customized test invoices. Review custom details before generating. Stored invoices can be searched by purpose, invoice number, batch ID, or date in the SharePoint Invoice Library."


async def _send_card(context: TurnContext, card: dict) -> None:
    logger.info("response_type=adaptive_card correlation_id=%s", context.activity.id or "unknown")
    await context.send_activity(Activity(type=ActivityTypes.message, attachments=[Attachment(contentType="application/vnd.microsoft.card.adaptive", content=card)]))


async def _send_text(context: TurnContext, message: str) -> None:
    logger.info("response_type=text correlation_id=%s", context.activity.id or "unknown")
    await context.send_activity(message)


def _log_incoming(context: TurnContext) -> None:
    text = (context.activity.text or "").replace("\r", " ").replace("\n", " ")[:100]
    text = re.sub(r"[\w.+-]+@[\w.-]+", "[email]", text)
    text = re.sub(r"\d", "*", text)
    value = context.activity.value
    action = (value.get("action") or value.get("route")) if isinstance(value, dict) else None
    if isinstance(value, dict) and action in REVIEW_ACTIONS:
        visible = ("invoice_type", "purpose", "date_from", "date_to", "min_amount", "max_amount", "currency", "tax_type", "tax_rate", "line_item_mode", "max_line_items", "company_mode", "count")
        logged_value = {"route": action, **{field: value.get(field) for field in visible if field in value}, "company_fields": "[redacted]"}
    else:
        logged_value = value if action in HOME_ACTIONS else {"action": action, "fields": sorted(value)} if isinstance(value, dict) else value
    logger.info("incoming_activity type=%s name=%s value=%s card_action=%s correlation_id=%s text_preview=%r", context.activity.type, context.activity.name, logged_value, action, context.activity.id or "unknown", text)


def _name(context: TurnContext) -> str:
    account = context.activity.from_property
    name = (account.name or "there") if account else "there"
    return " ".join(name.split())[:80] or "there"


def _user_id(context: TurnContext) -> str:
    account = context.activity.from_property
    return account.id if account and account.id else "local-user"


def _actor(context: TurnContext) -> GenerationActor:
    account = context.activity.from_property
    raw_channel_data = getattr(context.activity, "channel_data", None)
    channel_data = raw_channel_data if isinstance(raw_channel_data, dict) else {}
    email = getattr(account, "email", None) if account else None
    email = email or channel_data.get("userPrincipalName") or channel_data.get("userEmail")
    return GenerationActor(
        display_name=_name(context),
        aad_object_id=getattr(account, "aad_object_id", None) if account else None,
        email=email,
    )


def _active_pending(context: TurnContext, pending: object, token: object) -> bool:
    if not isinstance(pending, dict) or pending.get("token") != token or pending.get("user_id") != _user_id(context):
        return False
    try:
        created_at = datetime.fromisoformat(pending["created_at"])
    except (KeyError, TypeError, ValueError):
        return False
    return created_at.tzinfo is not None and datetime.now(timezone.utc) - created_at < timedelta(minutes=15)


async def _home(context: TurnContext) -> None:
    await _send_text(context, f"Hi {_name(context)}, welcome to the Invoice Generator.\nWhat would you like to do?")
    await _send_card(context, home_card())


@agent_app.conversation_update("membersAdded")
async def on_members_added(context: TurnContext, _state: TurnState) -> None:
    logger.info("incoming_activity type=membersAdded correlation_id=%s selected_route=home", context.activity.id or "unknown")
    await _home(context)


async def _handle_home_action(context: TurnContext, state: TurnState, action: str) -> None:
    """Handle Home card submissions before natural-language routing."""
    if action == "generate_invoices":
        state.delete_value(PENDING_KEY)
        state.delete_value(PREFILL_KEY)
        logger.info("generation_method_selected method=choose correlation_id=%s", context.activity.id or "unknown")
        await _send_card(context, generation_method_card())
    elif action == "search_invoices":
        await _send_text(context, "Ask me anything about your invoices.")
        await _send_card(context, search_prompt_card())
    elif action == "help":
        await _send_text(context, HELP_TEXT)


@agent_app.activity(ActivityTypes.message)
async def on_message(context: TurnContext, state: TurnState) -> None:
    _log_incoming(context)
    correlation_id = context.activity.id or uuid4().hex
    value = context.activity.value if isinstance(context.activity.value, dict) else {}
    card_action = value.get("action") or value.get("route")
    if card_action in HOME_ACTIONS:
        logger.info("selected_route=%s correlation_id=%s", card_action, correlation_id)
        try:
            await _handle_home_action(context, state, card_action)
        except Exception:
            logger.exception("Failed handling adaptive card action: %s", card_action)
            raise
        return
    route = route_message(context.activity.text, context.activity.value)
    logger.info("selected_route=%s correlation_id=%s", route, correlation_id)
    pending = state.get_value(PENDING_KEY, lambda: None)

    if route in {"home", "go_home"}:
        state.delete_value(PENDING_KEY)
        state.delete_value(PREFILL_KEY)
        await _home(context)
    elif route == "help":
        await _send_text(context, HELP_TEXT)
    elif route in {"search", "view_batch"}:
        try:
            if route == "view_batch":
                batch_id = str(value.get("batch_id") or "")
                if not re.fullmatch(r"BATCH-[0-9]{8}-[A-Z0-9]+", batch_id):
                    raise InvoiceValidationError("Choose a valid batch.")
                records = await asyncio.to_thread(sharepoint_service.list_batch_invoices, batch_id)
                title = f"Batch {batch_id}"
            elif not (context.activity.text or "").strip():
                await _send_text(context, "Ask me anything about your invoices.")
                await _send_card(context, search_prompt_card())
                return
            else:
                query = context.activity.text.strip()
                records = await asyncio.to_thread(sharepoint_service.search_invoice_metadata, query, 10)
                title = "Invoice Library results"
            await _send_card(context, search_results_card(records, title))
        except InvoiceValidationError as exc:
            await _send_text(context, str(exc))
        except SharePointError as exc:
            logger.exception("sharepoint_search_failed correlation_id=%s", correlation_id)
            await _send_text(context, f"I couldn't search the Invoice Library. {exc}")
    elif route == "generate":
        state.delete_value(PENDING_KEY)
        prefill = prefill_from_text(context.activity.text)
        count = requested_invoice_count(context.activity.text)
        if count and count > 1:
            prefill.update(count=str(count), generation_mode="Batch")
        state.set_value(PREFILL_KEY, prefill)
        logger.info("generation_method_selected method=choose correlation_id=%s", correlation_id)
        await _send_card(context, generation_method_card())
    elif route == "generate_random":
        logger.info("generation_method_selected method=random correlation_id=%s", correlation_id)
        await _send_card(context, random_generation_card(settings.max_batch_size, settings.default_purpose))
    elif route == "customize_invoice":
        logger.info("generation_method_selected method=customize correlation_id=%s", correlation_id)
        prefill = state.get_value(PREFILL_KEY, lambda: {})
        state.delete_value(PREFILL_KEY)
        await _send_card(context, invoice_form_card(prefill))
    elif route == "generate_random_batch":
        try:
            try:
                random_request = RandomGenerationRequest.from_form(value, settings.max_batch_size, settings.default_purpose)
            except InvoiceValidationError as exc:
                await _send_text(context, str(exc))
                await _send_card(context, random_generation_card(settings.max_batch_size, settings.default_purpose, str(exc), value))
                return
            logger.info("generation_method_selected method=random requested_count=%s correlation_id=%s", random_request.count, correlation_id)
            result = await asyncio.to_thread(generation_service.generate_random, random_request, _actor(context))
            await _send_card(context, result_card(result, sharepoint_service.library_url))
        except SharePointError as exc:
            logger.exception("random_generation_sharepoint_failed correlation_id=%s", correlation_id)
            await _send_text(context, f"No invoices were stored in the Invoice Library. {exc}")
        except Exception:
            logger.exception("random_generation_failed correlation_id=%s", correlation_id)
            raise
    elif route in REVIEW_ACTIONS:
        try:
            logger.info("submitted_form_fields=%s correlation_id=%s", sorted(value.keys()), correlation_id)
            try:
                request = InvoiceRequest.from_form({**value, "company_mode": "Custom"}, settings.max_batch_size, settings.max_line_items, settings.max_amount)
            except InvoiceValidationError as exc:
                await _send_text(context, str(exc))
                await _send_card(context, invoice_form_card(value, str(exc)))
                return
            token = uuid4().hex
            state.set_value(PENDING_KEY, {"token": token, "form": request.to_form(), "user_id": _user_id(context), "created_at": datetime.now(timezone.utc).isoformat()})
            await _send_text(context, review_summary_text(request))
            await _send_card(context, review_card(request, token))
        except Exception:
            logger.exception("Failed handling adaptive card action: review_invoice_request")
            raise
    elif route == "edit":
        if not _active_pending(context, pending, value.get("token")):
            await _send_text(context, "That review is no longer active. Please start a new invoice.")
            return
        state.delete_value(PENDING_KEY)
        await _send_card(context, invoice_form_card(pending["form"]))
    elif route == "cancel":
        state.delete_value(PENDING_KEY)
        await _send_text(context, "Invoice creation cancelled.")
        await _home(context)
    elif route in CONFIRM_ACTIONS:
        if not _active_pending(context, pending, value.get("token")):
            await _send_text(context, "That review is no longer active. Please start a new invoice.")
            return
        try:
            request = InvoiceRequest.from_form(pending["form"], settings.max_batch_size, settings.max_line_items, settings.max_amount)
            state.delete_value(PENDING_KEY)
            logger.info("generation_method_selected method=customize requested_count=%s correlation_id=%s", request.count, correlation_id)
            result = await asyncio.to_thread(generation_service.generate_custom, request, _actor(context))
        except (InvoiceValidationError, InvoiceGenerationError) as exc:
            logger.warning("generation_rejected correlation_id=%s reason=%s", correlation_id, exc)
            await _send_text(context, str(exc))
            return
        except SharePointError as exc:
            logger.exception("custom_generation_sharepoint_failed correlation_id=%s", correlation_id)
            await _send_text(context, f"No invoices were stored in the Invoice Library. {exc}")
            return
        except Exception:
            logger.exception("custom_generation_failed correlation_id=%s", correlation_id)
            raise
        logger.info("invoice_generation_completed correlation_id=%s batch_id=%s generated=%s stored=%s failed=%s", correlation_id, result.batch_id, result.generated_count, result.uploaded_count, result.failed_count)
        await _send_card(context, result_card(result, sharepoint_service.library_url))
    else:
        if context.activity.text:
            try:
                answer = await answer_general_question(context.activity.text)
            except Exception:
                logger.exception("llm_fallback_failed correlation_id=%s", correlation_id)
                answer = "I couldn't answer that right now. Choose Help to see what I can do."
            await _send_text(context, answer)
        else:
            await _send_text(context, "Choose Generate Invoices, Search Invoices, or Help to get started.")
            await _home(context)


@agent_app.error
async def on_error(context: TurnContext, error: Exception) -> None:
    logger.exception("agent_error correlation_id=%s", context.activity.id if context.activity else "unknown", exc_info=error)
    await _send_text(context, "Something went wrong while handling that request. Please try again.")
