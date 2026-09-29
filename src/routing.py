"""Explicit intent routing; unknown text never triggers invoice creation."""

from __future__ import annotations

import re
from datetime import date
from typing import Any, Mapping


_MONTHS = {name.lower(): index for index, name in enumerate(("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"), 1)}


def route_message(text: str | None, value: Any) -> str:
    if isinstance(value, Mapping) and value.get("route") in {
        "home", "go_home", "generate", "generate_invoices", "generate_random",
        "generate_random_batch", "customize_invoice", "search", "search_invoices",
        "view_batch", "help", "review", "review_invoice_request", "confirm",
        "generate_confirmed_invoices", "edit", "cancel",
    }:
        return str(value["route"])
    content = (text or "").strip().lower()
    if not content or content in {"hi", "hello", "start", "home", "menu"}:
        return "home"
    if content in {"help", "what can you do?"}:
        return "help"
    if content.startswith(("search", "find", "show", "list", "which", "compare", "what was")):
        return "search"
    if re.search(r"\bINV_[A-Z0-9_]+\b", text or "", re.IGNORECASE):
        return "search"
    if content.startswith(("generate", "create", "make")):
        return "generate"
    return "unknown"


def prefill_from_text(text: str | None) -> dict[str, str]:
    """Extract only simple explicit hints; the form remains the source of truth."""
    content = text or ""
    lower = content.lower()
    values: dict[str, str] = {}
    if re.search(r"\bgst\b|\btax invoices?\b", lower):
        values.update(invoice_type="GST", tax_type="IGST", tax_rate="18")
    for purpose in ("uat", "demo", "training", "development", "testing"):
        if re.search(rf"\b{purpose}\b", lower):
            values["purpose"] = purpose.upper() if purpose == "uat" else purpose.title()
            break
    for month, index in _MONTHS.items():
        if re.search(rf"\b{month}\b", lower):
            year_match = re.search(r"\b20\d{2}\b", lower)
            year = int(year_match.group()) if year_match else date.today().year
            if 1 <= year <= 9999:
                from calendar import monthrange
                values["date_from"] = date(year, index, 1).isoformat()
                values["date_to"] = date(year, index, monthrange(year, index)[1]).isoformat()
            break
    amount_match = re.search(r"between\s*(?:₹|inr\s*)?([\d,]+(?:\.\d{1,2})?)\s*(?:and|to|-)\s*(?:₹|inr\s*)?([\d,]+(?:\.\d{1,2})?)", lower)
    if amount_match:
        values["min_amount"] = amount_match.group(1).replace(",", "")
        values["max_amount"] = amount_match.group(2).replace(",", "")
        if "₹" in amount_match.group(0) or "inr" in amount_match.group(0):
            values["currency"] = "INR"
    return values


def requested_invoice_count(text: str | None) -> int | None:
    match = re.search(r"\b(?:generate|create|make)\s+(\d+)\s+(?:\w+\s+){0,2}?invoices?\b", text or "", re.IGNORECASE)
    return int(match.group(1)) if match else None
