"""Pure invoice generation and accounting calculations."""

from __future__ import annotations

import random
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, ROUND_HALF_UP

from domain import Invoice, InvoiceRequest, InvoiceTotals, LineItem, Party, TaxBreakdown
from errors import InvoiceGenerationError, InvoiceValidationError


CENT = Decimal("0.01")


def rounded(value: Decimal) -> Decimal:
    return value.quantize(CENT, rounding=ROUND_HALF_UP)


def calculate_totals(items: tuple[LineItem, ...], tax_type: str, rate: Decimal) -> InvoiceTotals:
    if not items or any(item.quantity < 1 or item.unit_price <= 0 or item.amount != rounded(item.unit_price * item.quantity) for item in items):
        raise InvoiceValidationError("Invoice line items must contain valid positive amounts.")
    subtotal = sum((item.amount for item in items), Decimal("0.00"))
    zero = Decimal("0.00")
    cgst = sgst = igst = zero
    if tax_type == "CGST/SGST":
        cgst = rounded(subtotal * rate / Decimal("200"))
        sgst = rounded(subtotal * rate / Decimal("200"))
    elif tax_type == "IGST":
        igst = rounded(subtotal * rate / Decimal("100"))
    elif tax_type != "None":
        raise InvoiceValidationError("Unsupported tax type.")
    tax = TaxBreakdown(tax_type, rate, cgst, sgst, igst)
    return InvoiceTotals(subtotal, zero, subtotal, tax, subtotal + cgst + sgst + igst)


def _base_for_range(request: InvoiceRequest, rng: random.Random) -> Decimal:
    low = int(request.min_amount * 100)
    high = int(request.max_amount * 100)
    desired = rng.randint(low, high)
    estimate = int(Decimal(desired) / (Decimal("1") + request.tax_rate / 100))
    for offset in (0, *range(1, 102), *range(-1, -102, -1)):
        cents = estimate + offset
        if cents <= 0:
            continue
        base = Decimal(cents) / 100
        probe = calculate_totals((LineItem("Item", 1, base, base),), request.tax_type, request.tax_rate)
        if request.min_amount <= probe.grand_total <= request.max_amount:
            return base
    raise InvoiceGenerationError("The amount range is too narrow for the selected tax rate. Widen it slightly and try again.")


_PRODUCTS = ("Laptop accessory", "Network equipment", "Office supplies", "Security device")
_SERVICES = ("Implementation service", "Consulting service", "Support service", "Training service")
_SELLERS = (
    Party("Northstar Systems Pvt Ltd", "12 Innovation Road, Bengaluru, Karnataka 560001", "29AAACN1234A1Z5"),
    Party("Harborline Solutions Pvt Ltd", "84 Market Street, Mumbai, Maharashtra 400001", "27AAACH5678B1Z2"),
) + tuple(
    Party(f"{prefix} {sector} Pvt Ltd", f"{index + 10} Test Commerce Road, Bengaluru, Karnataka 560001")
    for index, prefix in enumerate(("Aster", "Beacon", "Cedar", "Delta", "Evergreen", "Falcon", "Granite", "Horizon", "Indigo", "Juniper", "Keystone", "Lumen"))
    for sector in ("Systems", "Analytics", "Equipment", "Services")
)
_BUYERS = (
    Party("Apex Retail Pvt Ltd", "22 Commerce Avenue, Pune, Maharashtra 411001", "27AAAFA1234A1Z4"),
    Party("Meridian Labs Pvt Ltd", "7 Research Park, Hyderabad, Telangana 500001", "36AAACM5678C1Z6"),
) + tuple(
    Party(f"{prefix} {sector} Pvt Ltd", f"{index + 20} Demo Business Park, Pune, Maharashtra 411001")
    for index, prefix in enumerate(("Orion", "Pioneer", "Quartz", "Redwood", "Summit", "Terra", "Unity", "Vertex", "Willow", "Zenith", "Cobalt", "Mosaic"))
    for sector in ("Retail", "Laboratories", "Trading", "Operations")
)


def _line_items(base: Decimal, request: InvoiceRequest, rng: random.Random) -> tuple[LineItem, ...]:
    base_cents = int(base * 100)
    count = rng.randint(1, min(request.max_line_items, base_cents))
    if count == 1:
        parts = [base_cents]
    else:
        cuts = sorted(rng.sample(range(1, base_cents), count - 1))
        points = [0, *cuts, base_cents]
        parts = [points[index + 1] - points[index] for index in range(count)]
    descriptions = _PRODUCTS if request.line_item_mode == "Products" else _SERVICES if request.line_item_mode == "Services" else _PRODUCTS + _SERVICES
    return tuple(
        LineItem(rng.choice(descriptions), 1, Decimal(cents) / 100, Decimal(cents) / 100)
        for cents in parts
    )


def _custom_line_items(base: Decimal, request: InvoiceRequest) -> tuple[LineItem, ...]:
    """Keep the description and structure stable across a customized batch."""
    minimum_base_cents = max(1, int(request.min_amount * 100 / (Decimal("1") + request.tax_rate / 100)) - 2)
    count = min(request.max_line_items, minimum_base_cents)
    base_cents = int(base * 100)
    if base_cents < count:
        raise InvoiceGenerationError("The amount range is too narrow for the selected line item count.")
    descriptions = _PRODUCTS if request.line_item_mode == "Products" else _SERVICES if request.line_item_mode == "Services" else _PRODUCTS + _SERVICES
    whole, remainder = divmod(base_cents, count)
    return tuple(
        LineItem(descriptions[index % len(descriptions)], 1, Decimal(whole + (index < remainder)) / 100, Decimal(whole + (index < remainder)) / 100)
        for index in range(count)
    )


def generate_invoice(
    request: InvoiceRequest,
    invoice_number: str,
    batch_id: str,
    generated_by: str,
    rng: random.Random | None = None,
    *,
    fixed_details: bool = False,
    fixed_date: date | None = None,
) -> Invoice:
    """Generate one invoice. IDs and persistence are supplied by the caller."""
    rng = rng or random.Random()
    days = (request.date_to - request.date_from).days
    invoice_date = fixed_date or request.date_from + timedelta(days=rng.randint(0, days))
    base = _base_for_range(request, rng)
    items = _custom_line_items(base, request) if fixed_details else _line_items(base, request, rng)
    totals = calculate_totals(items, request.tax_type, request.tax_rate)
    seller = Party(request.seller_name, request.seller_address, request.seller_tax_id) if request.company_mode == "Custom" else rng.choice(_SELLERS)
    buyer = Party(request.buyer_name, request.buyer_address, request.buyer_tax_id) if request.company_mode == "Custom" else rng.choice(_BUYERS)
    invoice = Invoice(
        invoice_number=invoice_number,
        invoice_type=request.invoice_type,
        purpose=request.purpose,
        invoice_date=invoice_date,
        due_date=invoice_date + timedelta(days=30),
        currency=request.currency,
        seller=seller,
        buyer=buyer,
        line_items=items,
        totals=totals,
        batch_id=batch_id,
        generated_at=datetime.now(timezone.utc),
        generated_by=generated_by,
    )
    validate_invoice(invoice, request)
    return invoice


def validate_invoice(invoice: Invoice, request: InvoiceRequest) -> None:
    recomputed = calculate_totals(invoice.line_items, request.tax_type, request.tax_rate)
    if invoice.totals != recomputed:
        raise InvoiceValidationError("Invoice totals do not reconcile.")
    if not request.min_amount <= invoice.totals.grand_total <= request.max_amount:
        raise InvoiceValidationError("Invoice total is outside the requested range.")
    if not request.date_from <= invoice.invoice_date <= request.date_to:
        raise InvoiceValidationError("Invoice date is outside the requested range.")
    if not invoice.seller.name or not invoice.buyer.name:
        raise InvoiceValidationError("Seller and buyer names are required.")
