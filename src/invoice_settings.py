"""Environment-backed application settings."""

from __future__ import annotations

import os
import json
import tempfile
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Mapping

from dotenv import load_dotenv


@dataclass(frozen=True)
class Settings:
    host: str
    port: int
    output_dir: Path
    max_batch_size: int
    max_line_items: int
    max_amount: Decimal
    log_level: str
    azure_openai_api_key: str | None
    azure_openai_deployment_name: str | None
    azure_openai_endpoint: str | None
    sharepoint_site_id: str | None
    sharepoint_library_id: str | None
    sharepoint_library_name: str
    graph_tenant_id: str | None
    graph_client_id: str | None
    graph_client_secret: str | None
    user_lookup_ids: dict[str, int]
    default_purpose: str
    default_currency: str
    random_generation_max_retries: int
    random_min_amount: Decimal
    random_max_amount: Decimal
    search_max_items: int

    @classmethod
    def from_env(cls, env: Mapping[str, str] | None = None) -> "Settings":
        if env is None:
            load_dotenv()
            env = os.environ
        port = int(env.get("PORT") or "3978")
        max_batch_size = int(env.get("INVOICE_MAX_BATCH_SIZE") or "50")
        max_line_items = int(env.get("INVOICE_MAX_LINE_ITEMS") or "20")
        try:
            max_amount = Decimal(env.get("INVOICE_MAX_AMOUNT") or "100000000")
            random_min_amount = Decimal(env.get("RANDOM_MIN_AMOUNT") or "1000")
            random_max_amount = Decimal(env.get("RANDOM_MAX_AMOUNT") or "10000")
        except InvalidOperation:
            raise ValueError("Invoice amount settings must be valid decimals") from None
        retries = int(env.get("RANDOM_GENERATION_MAX_RETRIES") or "30")
        search_max_items = int(env.get("SHAREPOINT_SEARCH_MAX_ITEMS") or "5000")
        try:
            raw_lookup_ids = json.loads(env.get("SHAREPOINT_USER_LOOKUP_IDS") or "{}")
            if not isinstance(raw_lookup_ids, dict):
                raise ValueError
            user_lookup_ids = {}
            for key, value in raw_lookup_ids.items():
                if not str(key).strip() or isinstance(value, bool) or not str(value).isdigit():
                    raise ValueError
                user_lookup_ids[str(key).casefold()] = int(value)
        except (ValueError, TypeError, AttributeError):
            raise ValueError("SHAREPOINT_USER_LOOKUP_IDS must be a JSON object of positive lookup IDs") from None
        if not 1 <= port <= 65535 or max_batch_size < 1 or max_line_items < 1 or retries < 1 or search_max_items < 1 or not max_amount.is_finite() or not 0 < max_amount <= Decimal("1000000000000"):
            raise ValueError("An invoice or hosting setting is outside its valid range")
        if not random_min_amount.is_finite() or not random_max_amount.is_finite() or not Decimal("0") < random_min_amount <= random_max_amount <= max_amount:
            raise ValueError("Random amount range is outside the configured invoice amount range")
        if any(value < 1 for value in user_lookup_ids.values()):
            raise ValueError("SharePoint person lookup IDs must be positive")
        from domain import CURRENCIES, PURPOSES
        default_purpose = env.get("DEFAULT_PURPOSE") or "Testing"
        default_currency = env.get("DEFAULT_CURRENCY") or "INR"
        if default_purpose not in PURPOSES or default_currency not in CURRENCIES:
            raise ValueError("DEFAULT_PURPOSE or DEFAULT_CURRENCY is unsupported")
        log_level = (env.get("LOG_LEVEL") or "INFO").upper()
        if log_level not in {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}:
            raise ValueError("LOG_LEVEL must be a standard Python logging level")
        output_dir = Path(env.get("INVOICE_OUTPUT_DIR") or str(Path(tempfile.gettempdir()) / "invoice-generator"))
        if not output_dir.is_absolute():
            output_dir = Path(__file__).resolve().parent.parent / output_dir
        return cls(
            host=env.get("HOST") or "localhost",
            port=port,
            output_dir=output_dir,
            max_batch_size=max_batch_size,
            max_line_items=max_line_items,
            max_amount=max_amount,
            log_level=log_level,
            azure_openai_api_key=env.get("AZURE_OPENAI_API_KEY") or None,
            azure_openai_deployment_name=env.get("AZURE_OPENAI_DEPLOYMENT_NAME") or None,
            azure_openai_endpoint=env.get("AZURE_OPENAI_ENDPOINT") or None,
            sharepoint_site_id=env.get("SHAREPOINT_SITE_ID") or None,
            sharepoint_library_id=env.get("SHAREPOINT_LIBRARY_ID") or None,
            sharepoint_library_name=env.get("SHAREPOINT_LIBRARY_NAME") or "Invoice Library",
            graph_tenant_id=env.get("SHAREPOINT_GRAPH_TENANT_ID") or env.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID") or None,
            graph_client_id=env.get("SHAREPOINT_GRAPH_CLIENT_ID") or env.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID") or None,
            graph_client_secret=env.get("SHAREPOINT_GRAPH_CLIENT_SECRET") or env.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET") or None,
            user_lookup_ids=user_lookup_ids,
            default_purpose=default_purpose,
            default_currency=default_currency,
            random_generation_max_retries=retries,
            random_min_amount=random_min_amount,
            random_max_amount=random_max_amount,
            search_max_items=search_max_items,
        )


settings = Settings.from_env()
