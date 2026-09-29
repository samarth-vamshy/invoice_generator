"""Adaptive Card views for the local invoice workflow."""

from __future__ import annotations

from datetime import date
from typing import Any, Mapping

from domain import BatchResult, CURRENCIES, INVOICE_TYPES, LINE_ITEM_MODES, PURPOSES, TAX_TYPES, InvoiceRequest
from sharepoint_service import SearchRecord


def _card(body: list[dict[str, Any]], actions: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    return {"type": "AdaptiveCard", "version": "1.5", "body": body, "actions": actions or [], "$schema": "http://adaptivecards.io/schemas/adaptive-card.json"}


def _text(text: str, **kwargs: Any) -> dict[str, Any]:
    return {"type": "TextBlock", "text": text, "wrap": True, **kwargs}


def _submit(title: str, route: str, style: str | None = None, **data: str) -> dict[str, Any]:
    inputs = "auto" if route in {"review_invoice_request", "generate_random_batch"} else "none"
    action = {"type": "Action.Execute", "title": title, "verb": route, "data": {"action": route, **data}, "associatedInputs": inputs}
    if style:
        action["style"] = style
    return action


def home_card() -> dict[str, Any]:
    return _card([_text("Invoice Generator", size="Large", weight="Bolder"), _text("Choose an action to continue.")], [
        _submit("Generate Invoices", "generate_invoices"), _submit("Search Invoices", "search_invoices"), _submit("Help", "help")
    ])


def generation_method_card() -> dict[str, Any]:
    return _card(
        [_text("Generate Invoices", size="Large", weight="Bolder"), _text("How would you like to proceed?")],
        [_submit("Generate Random", "generate_random"), _submit("Customize", "customize_invoice"), _submit("Home", "go_home")],
    )


def random_generation_card(max_batch_size: int, default_purpose: str, error: str | None = None, values: Mapping[str, Any] | None = None) -> dict[str, Any]:
    values = values or {}
    try:
        count_value = int(str(values.get("invoice_count") or "1"))
    except (TypeError, ValueError):
        count_value = 1
    if not 1 <= count_value <= max_batch_size:
        count_value = 1
    body = [_text("Generate Random Invoices", size="Large", weight="Bolder"), _text("The PDFs will be stored in the Invoice Library.")]
    if error:
        body.append(_text(error, color="Attention"))
    body += [
        {"type": "Input.Number", "id": "invoice_count", "label": "Number of invoices", "value": count_value, "min": 1, "max": max_batch_size, "isRequired": True, "errorMessage": f"Enter a whole number from 1 to {max_batch_size}."},
        _choice("purpose", "Purpose", PURPOSES, str(values.get("purpose") or default_purpose)),
    ]
    return _card(body, [_submit("Generate", "generate_random_batch", "positive"), _submit("Home", "go_home")])


def _choice(id: str, label: str, values: tuple[str, ...], selected: str) -> dict[str, Any]:
    return {"type": "Input.ChoiceSet", "id": id, "label": label, "style": "compact", "isRequired": True, "errorMessage": f"Choose {label.lower()}.", "value": selected, "choices": [{"title": value, "value": value} for value in values]}


def _input(id: str, label: str, value: str, required: bool = True, **kwargs: Any) -> dict[str, Any]:
    field = {"type": "Input.Text", "id": id, "label": label, "value": value, "isRequired": required, **kwargs}
    if required:
        field["errorMessage"] = f"Enter {label.lower()}."
    return field


def invoice_form_card(values: Mapping[str, Any] | None = None, error: str | None = None) -> dict[str, Any]:
    values = values or {}
    get = lambda key, default: str(values.get(key) or default)
    today = date.today().isoformat()
    body = [_text("Generate an invoice", size="Large", weight="Bolder"), _text("Enter your seller, buyer, and tax details. A batch keeps those details fixed while invoice amounts vary within your range.", isSubtle=True)]
    if error:
        body.append(_text(error, color="Attention"))
    body += [
        _choice("invoice_type", "Invoice type", INVOICE_TYPES, get("invoice_type", "Standard")),
        _choice("purpose", "Purpose", PURPOSES, get("purpose", "Testing")),
        _choice("generation_mode", "Generation mode", ("Single", "Batch"), get("generation_mode", "Single")),
        _input("count", "Invoice count", get("count", "1")),
        {"type": "Input.Date", "id": "date_from", "label": "Start date", "value": get("date_from", today), "isRequired": True, "errorMessage": "Choose a start date."},
        {"type": "Input.Date", "id": "date_to", "label": "End date", "value": get("date_to", today), "isRequired": True, "errorMessage": "Choose an end date."},
        _input("min_amount", "Minimum total amount", get("min_amount", "1000.00")),
        _input("max_amount", "Maximum total amount", get("max_amount", "10000.00")),
        _choice("currency", "Currency", CURRENCIES, get("currency", "INR")),
        _choice("tax_type", "Tax", TAX_TYPES, get("tax_type", "None")),
        _input("tax_rate", "Tax rate (%) — enter 0 for no tax", get("tax_rate", "0")),
        _choice("line_item_mode", "Line items", LINE_ITEM_MODES, get("line_item_mode", "Mixed")),
        _input("max_line_items", "Maximum line items", get("max_line_items", "3")),
        _text("Seller and buyer names are required. Addresses and tax IDs are used exactly as entered.", isSubtle=True),
        _input("seller_name", "Seller name", get("seller_name", "")),
        _input("seller_address", "Seller address", get("seller_address", ""), False),
        _input("seller_tax_id", "Seller tax ID", get("seller_tax_id", ""), False),
        _input("buyer_name", "Buyer name", get("buyer_name", "")),
        _input("buyer_address", "Buyer address", get("buyer_address", ""), False),
        _input("buyer_tax_id", "Buyer tax ID", get("buyer_tax_id", ""), False),
        _choice("output_destination", "Output", ("PDF in Invoice Library",), get("output_destination", "PDF in Invoice Library")),
    ]
    return _card(body, [_submit("Review", "review_invoice_request", company_mode="Custom"), _submit("Cancel", "cancel"), _submit("Home", "go_home")])


def _review_summary(request: InvoiceRequest) -> list[tuple[str, str]]:
    return [
        ("Invoice type", request.invoice_type), ("Purpose", request.purpose), ("Generation mode", request.generation_mode),
        ("Invoice count", str(request.count)),
        ("Date range", f"{request.date_from:%d %b %Y} – {request.date_to:%d %b %Y}"),
        ("Amount range", f"{request.currency} {request.min_amount:,.2f} – {request.max_amount:,.2f}"),
        ("Currency", request.currency),
        ("Tax type", request.tax_type), ("Tax rate", f"{request.tax_rate.normalize():g}%"),
        ("Line item mode", request.line_item_mode),
        ("Seller name", request.seller_name), ("Seller address", request.seller_address or "Not provided"),
        ("Seller tax ID", request.seller_tax_id or "Not provided"),
        ("Buyer name", request.buyer_name), ("Buyer address", request.buyer_address or "Not provided"),
        ("Buyer tax ID", request.buyer_tax_id or "Not provided"),
        ("Output", request.output_destination),
    ]


def review_summary_text(request: InvoiceRequest) -> str:
    details = "\n".join(f"{name}: {value}" for name, value in _review_summary(request))
    return f"Review your invoice:\n{details}\nNo document has been created. Choose Generate, Edit, or Cancel on the card below."


def review_card(request: InvoiceRequest, token: str) -> dict[str, Any]:
    body = [_text("Review your invoice", size="Large", weight="Bolder"), _text("No document is created until you choose Generate. A batch keeps these details and its chosen invoice date fixed; amounts and invoice numbers vary.")]
    body += [{"type": "FactSet", "facts": [{"title": name, "value": value} for name, value in _review_summary(request)]}]
    return _card(body, [_submit("Generate", "generate_confirmed_invoices", "positive", token=token), _submit("Edit", "edit", token=token), _submit("Cancel", "cancel"), _submit("Home", "go_home")])


def result_card(result: BatchResult, library_url: str) -> dict[str, Any]:
    complete = result.uploaded_count == result.requested_count
    if result.requested_count == 1 and complete:
        record = result.sharepoint_records[0]
        invoice = record.invoice
        body = [_text("Invoice generated successfully and stored in the Invoice Library.", weight="Bolder", color="Good")]
        body += [{"type": "FactSet", "facts": [
            {"title": "Invoice Number", "value": invoice.invoice_number},
            {"title": "Purpose", "value": invoice.purpose},
            {"title": "Invoice Type", "value": invoice.invoice_type},
            {"title": "Total Amount", "value": f"{invoice.currency} {invoice.totals.grand_total:,.2f}"},
        ]}]
        actions = [{"type": "Action.OpenUrl", "title": "Open Invoice", "url": record.web_url}]
    else:
        title = (
            f"{result.uploaded_count} invoices were generated successfully and stored in the Invoice Library."
            if complete else "Invoice generation finished with storage failures."
        )
        body = [_text(title, weight="Bolder", color="Good" if complete else "Attention")]
        body += [{"type": "FactSet", "facts": [
            {"title": "Batch ID", "value": result.batch_id},
            {"title": "Requested", "value": str(result.requested_count)},
            {"title": "Generated", "value": str(result.generated_count)},
            {"title": "Stored", "value": str(result.uploaded_count)},
            {"title": "Failures", "value": str(result.failed_count)},
        ]}]
        if result.failed_items:
            body.append(_text("Failed items: " + ", ".join(str(item.index) for item in result.failed_items), color="Attention"))
        actions = [_submit("View Batch", "view_batch", batch_id=result.batch_id)] if result.uploaded_count else []
        if result.uploaded_count:
            actions.append({"type": "Action.OpenUrl", "title": "Open Invoice Library", "url": library_url})
    actions += [_submit("Generate More", "generate_invoices"), _submit("Home", "go_home")]
    return _card(body, actions)


def search_prompt_card() -> dict[str, Any]:
    return _card([_text("Search Invoices", size="Large", weight="Bolder"), _text("Ask me anything about your invoices, such as a purpose, invoice number, batch ID, or invoices generated today.")], [_submit("Home", "go_home")])


def search_results_card(records: list[SearchRecord], title: str) -> dict[str, Any]:
    body = [_text(title, size="Large", weight="Bolder")]
    for record in records:
        metadata = record.metadata
        number = str(metadata.get("Invoice Number") or "Invoice")
        body.append(_text(f"{number} | {metadata.get('Purpose') or ''} | {metadata.get('Total Amount') or ''}"))
        if record.web_url:
            body.append({"type": "ActionSet", "actions": [{"type": "Action.OpenUrl", "title": "Open Invoice", "url": record.web_url}]})
    if not records:
        body.append(_text("No matching invoices were found in the Invoice Library."))
    return _card(body, [_submit("Home", "go_home")])
