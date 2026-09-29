"""Invoice Generator routes on the Teams SDK app already hosted by this project."""

from __future__ import annotations

import asyncio
import logging
import re
from datetime import datetime, timedelta, timezone
from typing import Any, Awaitable, Callable
from uuid import uuid4

from microsoft_teams.api import (
    AdaptiveCardActionMessageResponse,
    AdaptiveCardInvokeActivity,
    AdaptiveCardInvokeResponse,
    ConversationUpdateActivity,
    MessageActivity,
)
from microsoft_teams.apps import ActivityContext, App
from microsoft_teams.cards import AdaptiveCard

from cards import (
    generation_method_card,
    home_card,
    invoice_form_card,
    random_generation_card,
    result_card,
    review_card,
    review_summary_text,
    search_prompt_card,
    search_results_card,
)
from domain import InvoiceRequest, RandomGenerationRequest
from errors import InvoiceGenerationError, InvoiceValidationError, SharePointError
from invoice_service import InvoiceGenerationService
from invoice_settings import Settings, settings
from routing import prefill_from_text, requested_invoice_count, route_message
from sharepoint_service import GenerationActor, SharePointInvoiceService


logger = logging.getLogger(__name__)
logging.basicConfig(level=getattr(logging, settings.log_level), format="%(asctime)s %(levelname)s %(name)s %(message)s")

HELP_TEXT = (
    "Choose Generate Invoices for random or customized test invoices. Review custom details "
    "before generating. Stored invoices can be searched by purpose, invoice number, batch ID, "
    "or date in the SharePoint Invoice Library."
)
ACTION_ROUTES = {
    "home", "go_home", "generate", "generate_invoices", "generate_random",
    "generate_random_batch", "customize_invoice", "search", "search_invoices",
    "view_batch", "help", "review", "review_invoice_request", "confirm",
    "generate_confirmed_invoices", "edit", "cancel",
}
Fallback = Callable[[Any, ActivityContext[MessageActivity]], Awaitable[None]]


class InvoiceTeamsRouter:
    def __init__(
        self,
        model: Any,
        fallback: Fallback,
        settings_override: Settings | None = None,
        sharepoint: SharePointInvoiceService | None = None,
        generation: InvoiceGenerationService | None = None,
    ) -> None:
        self.model = model
        self.fallback = fallback
        self.settings = settings_override or settings
        self.sharepoint = sharepoint or SharePointInvoiceService(self.settings)
        self.generation = generation or InvoiceGenerationService(self.settings, self.sharepoint)
        self.pending: dict[tuple[str, str], dict[str, Any]] = {}
        self.prefill: dict[tuple[str, str], dict[str, str]] = {}
        self.search_waiting: set[tuple[str, str]] = set()

    @staticmethod
    def _key(ctx: ActivityContext[Any]) -> tuple[str, str]:
        activity = ctx.activity
        return (activity.conversation.id, activity.from_.id)

    @staticmethod
    def _name(ctx: ActivityContext[Any]) -> str:
        name = (getattr(ctx.activity.from_, "name", None) or "there").strip()
        return " ".join(name.split())[:80] or "there"

    def _actor(self, ctx: ActivityContext[Any]) -> GenerationActor:
        account = ctx.activity.from_
        properties = account.properties or {}
        channel_data = ctx.activity.channel_data
        if hasattr(channel_data, "model_dump"):
            channel_data = channel_data.model_dump(by_alias=True, exclude_none=True)
        if not isinstance(channel_data, dict):
            channel_data = {}
        return GenerationActor(
            display_name=self._name(ctx),
            aad_object_id=account.aad_object_id,
            email=(
                getattr(account, "email", None)
                or properties.get("email")
                or properties.get("userPrincipalName")
                or channel_data.get("userPrincipalName")
                or channel_data.get("userEmail")
            ),
        )

    @staticmethod
    def _log_incoming(ctx: ActivityContext[Any], data: dict[str, Any], action: str | None) -> None:
        preview = re.sub(r"[\w.+-]+@[\w.-]+", "[email]", getattr(ctx.activity, "text", None) or "")
        preview = re.sub(r"\d", "*", preview.replace("\r", " ").replace("\n", " "))[:100]
        value: dict[str, Any] = {"action": action, "fields": sorted(data)}
        if action in {"review", "review_invoice_request"}:
            visible = (
                "invoice_type", "purpose", "generation_mode", "count", "date_from", "date_to",
                "min_amount", "max_amount", "currency", "tax_type", "tax_rate", "line_item_mode",
                "max_line_items", "company_mode",
            )
            value.update({field: data[field] for field in visible if field in data})
            value["company_fields"] = "[redacted]"
        logger.info(
            "incoming_activity type=%s name=%s value=%s action=%s verb=%s correlation_id=%s text_preview=%r",
            ctx.activity.type,
            getattr(ctx.activity, "name", None),
            value,
            action,
            getattr(getattr(getattr(ctx.activity, "value", None), "action", None), "verb", None),
            ctx.activity.id or "unknown",
            preview,
        )

    @staticmethod
    async def _send_text(ctx: ActivityContext[Any], message: str) -> None:
        logger.info("response_type=text correlation_id=%s", ctx.activity.id or "unknown")
        await ctx.send(message)

    @staticmethod
    async def _send_card(ctx: ActivityContext[Any], card: dict[str, Any]) -> None:
        logger.info("response_type=adaptive_card correlation_id=%s", ctx.activity.id or "unknown")
        await ctx.send(AdaptiveCard.model_validate(card))

    async def _home(self, ctx: ActivityContext[Any]) -> None:
        await self._send_text(ctx, f"Hi {self._name(ctx)}, welcome to the Invoice Generator.\nWhat would you like to do?")
        await self._send_card(ctx, home_card())

    def _active_pending(self, ctx: ActivityContext[Any], token: Any) -> dict[str, Any] | None:
        pending = self.pending.get(self._key(ctx))
        if not pending or not token or pending.get("token") != token:
            return None
        created_at = pending.get("created_at")
        if not isinstance(created_at, datetime) or datetime.now(timezone.utc) - created_at >= timedelta(minutes=15):
            self.pending.pop(self._key(ctx), None)
            return None
        return pending

    async def handle_conversation_update(self, ctx: ActivityContext[ConversationUpdateActivity]) -> None:
        activity = ctx.activity
        logger.info("incoming_activity type=%s name=%s correlation_id=%s", activity.type, getattr(activity, "name", None), activity.id or "unknown")
        if activity.members_added and activity.conversation.conversation_type == "personal":
            logger.info("selected_route=home correlation_id=%s", activity.id or "unknown")
            await self._home(ctx)

    async def handle_message(self, ctx: ActivityContext[MessageActivity]) -> None:
        activity = ctx.activity
        data = activity.value if isinstance(activity.value, dict) else {}
        action = data.get("action") or data.get("route")
        self._log_incoming(ctx, data, action)
        text = re.sub(r"<at>.*?</at>", "", activity.text or "", flags=re.IGNORECASE).strip()
        route = action if action in ACTION_ROUTES else route_message(text, data)
        if route == "unknown" and self._key(ctx) in self.search_waiting:
            route = "search"
        logger.info("selected_route=%s correlation_id=%s", route, activity.id or "unknown")
        try:
            await self.dispatch(ctx, route, data, text)
        except Exception:
            logger.exception("Failed handling Teams message route=%s correlation_id=%s", route, activity.id or "unknown")
            await self._send_text(ctx, "Something went wrong while handling that request. Please try again.")

    async def handle_card_action(self, ctx: ActivityContext[AdaptiveCardInvokeActivity]) -> AdaptiveCardInvokeResponse:
        activity = ctx.activity
        action = activity.value.action
        data = action.data or {}
        route = data.get("action") or data.get("route") or action.verb
        self._log_incoming(ctx, data, route)
        logger.info("selected_route=%s correlation_id=%s", route, activity.id or "unknown")
        try:
            if route not in ACTION_ROUTES:
                await self._send_text(ctx, "That card action is not recognized. Choose Home to start again.")
            else:
                await self.dispatch(ctx, route, data, "")
            return AdaptiveCardActionMessageResponse(value="Action processed")
        except Exception:
            logger.exception("Failed handling adaptive card action route=%s correlation_id=%s", route, activity.id or "unknown")
            await self._send_text(ctx, "Something went wrong while handling that request. Please try again.")
            return AdaptiveCardActionMessageResponse(value="The action could not be completed.")

    async def dispatch(self, ctx: ActivityContext[Any], route: str, data: dict[str, Any], text: str) -> None:
        key = self._key(ctx)
        correlation_id = ctx.activity.id or "unknown"
        if route in {"home", "go_home"}:
            self.pending.pop(key, None)
            self.prefill.pop(key, None)
            self.search_waiting.discard(key)
            await self._home(ctx)
        elif route == "help":
            await self._send_text(ctx, HELP_TEXT)
        elif route in {"generate", "generate_invoices"}:
            self.pending.pop(key, None)
            self.search_waiting.discard(key)
            prefill = prefill_from_text(text)
            count = requested_invoice_count(text)
            if count and count > 1:
                prefill.update(count=str(count), generation_mode="Batch")
            self.prefill[key] = prefill
            await self._send_card(ctx, generation_method_card())
        elif route == "generate_random":
            await self._send_card(ctx, random_generation_card(self.settings.max_batch_size, self.settings.default_purpose))
        elif route == "customize_invoice":
            await self._send_card(ctx, invoice_form_card(self.prefill.pop(key, {})))
        elif route == "generate_random_batch":
            try:
                request = RandomGenerationRequest.from_form(data, self.settings.max_batch_size, self.settings.default_purpose)
            except InvoiceValidationError as exc:
                await self._send_text(ctx, str(exc))
                await self._send_card(ctx, random_generation_card(self.settings.max_batch_size, self.settings.default_purpose, str(exc), data))
                return
            try:
                result = await asyncio.to_thread(self.generation.generate_random, request, self._actor(ctx))
                await self._send_card(ctx, result_card(result, self.sharepoint.library_url))
            except SharePointError as exc:
                logger.exception("random_generation_sharepoint_failed correlation_id=%s", correlation_id)
                await self._send_text(ctx, f"No invoices were stored in the Invoice Library. {exc}")
            except (InvoiceValidationError, InvoiceGenerationError) as exc:
                logger.exception("random_generation_rejected correlation_id=%s", correlation_id)
                await self._send_text(ctx, str(exc))
        elif route in {"review", "review_invoice_request"}:
            logger.info("submitted_form_fields=%s correlation_id=%s", sorted(data), correlation_id)
            try:
                request = InvoiceRequest.from_form(
                    {**data, "company_mode": "Custom"},
                    self.settings.max_batch_size,
                    self.settings.max_line_items,
                    self.settings.max_amount,
                )
            except InvoiceValidationError as exc:
                await self._send_text(ctx, str(exc))
                await self._send_card(ctx, invoice_form_card(data, str(exc)))
                return
            token = uuid4().hex
            self.pending[key] = {"token": token, "form": request.to_form(), "created_at": datetime.now(timezone.utc)}
            await self._send_text(ctx, review_summary_text(request))
            await self._send_card(ctx, review_card(request, token))
        elif route == "edit":
            pending = self._active_pending(ctx, data.get("token"))
            if not pending:
                await self._send_text(ctx, "That review is no longer active. Please start a new invoice.")
                return
            self.pending.pop(key, None)
            await self._send_card(ctx, invoice_form_card(pending["form"]))
        elif route == "cancel":
            self.pending.pop(key, None)
            self.search_waiting.discard(key)
            await self._send_text(ctx, "Invoice creation cancelled.")
            await self._home(ctx)
        elif route in {"confirm", "generate_confirmed_invoices"}:
            pending = self._active_pending(ctx, data.get("token"))
            if not pending:
                await self._send_text(ctx, "That review is no longer active. Please start a new invoice.")
                return
            request = InvoiceRequest.from_form(pending["form"], self.settings.max_batch_size, self.settings.max_line_items, self.settings.max_amount)
            self.pending.pop(key, None)
            try:
                result = await asyncio.to_thread(self.generation.generate_custom, request, self._actor(ctx))
                await self._send_card(ctx, result_card(result, self.sharepoint.library_url))
            except SharePointError as exc:
                logger.exception("custom_generation_sharepoint_failed correlation_id=%s", correlation_id)
                await self._send_text(ctx, f"No invoices were stored in the Invoice Library. {exc}")
            except (InvoiceValidationError, InvoiceGenerationError) as exc:
                logger.exception("custom_generation_rejected correlation_id=%s", correlation_id)
                await self._send_text(ctx, str(exc))
        elif route in {"search", "search_invoices", "view_batch"}:
            if route == "search_invoices" or (route == "search" and not text):
                self.search_waiting.add(key)
                await self._send_text(ctx, "Ask me anything about your invoices.")
                await self._send_card(ctx, search_prompt_card())
                return
            self.search_waiting.discard(key)
            try:
                if route == "view_batch":
                    batch_id = str(data.get("batch_id") or "")
                    if not re.fullmatch(r"BATCH-[0-9]{8}-[A-Z0-9]+", batch_id):
                        raise InvoiceValidationError("Choose a valid batch.")
                    records = await asyncio.to_thread(self.sharepoint.list_batch_invoices, batch_id)
                    title = f"Batch {batch_id}"
                else:
                    records = await asyncio.to_thread(self.sharepoint.search_invoice_metadata, text, 10)
                    title = "Invoice Library results"
                await self._send_card(ctx, search_results_card(records, title))
            except InvoiceValidationError as exc:
                await self._send_text(ctx, str(exc))
            except SharePointError as exc:
                logger.exception("sharepoint_search_failed correlation_id=%s", correlation_id)
                await self._send_text(ctx, f"I couldn't search the Invoice Library. {exc}")
        else:
            if text:
                try:
                    await self.fallback(self.model, ctx)
                except Exception:
                    logger.exception("llm_fallback_failed correlation_id=%s", correlation_id)
                    await self._send_text(ctx, "I couldn't answer that right now. Choose Help to see what I can do.")
            else:
                await self._home(ctx)


def register_invoice_routes(app: App, model: Any, fallback: Fallback) -> InvoiceTeamsRouter:
    router = InvoiceTeamsRouter(model, fallback)

    @app.on_message
    async def invoice_message(ctx: ActivityContext[MessageActivity]) -> None:
        await router.handle_message(ctx)

    @app.on_card_action
    async def invoice_card_action(ctx: ActivityContext[AdaptiveCardInvokeActivity]) -> AdaptiveCardInvokeResponse:
        return await router.handle_card_action(ctx)

    @app.on_conversation_update
    async def invoice_conversation_update(ctx: ActivityContext[ConversationUpdateActivity]) -> None:
        await router.handle_conversation_update(ctx)

    return router
