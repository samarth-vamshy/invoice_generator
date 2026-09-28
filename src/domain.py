"""Typed, serializable invoice domain objects."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any, Mapping

from errors import InvoiceValidationError


INVOICE_TYPES = ("Standard", "GST", "Service", "Product", "Proforma", "Commercial")
PURPOSES = ("Testing", "UAT", "Demo", "Training", "Development", "Other")
TAX_TYPES = ("None", "CGST/SGST", "IGST")
LINE_ITEM_MODES = ("Products", "Services", "Mixed")
CURRENCIES = ("INR", "USD", "EUR", "GBP")


def money(value: Any, label: str) -> Decimal:
    try:
        result = Decimal(str(value).strip().replace(",", ""))
    except (InvalidOperation, TypeError, ValueError):
        raise InvoiceValidationError(f"{label} must be a valid amount.") from None
    if not result.is_finite() or abs(result) > Decimal("1000000000000"):
        raise InvoiceValidationError(f"{label} is outside the supported amount range.")
    if result.as_tuple().exponent < -2:
        raise InvoiceValidationError(f"{label} must have at most two decimal places.")
    return result.quantize(Decimal("0.01"))


def _date(value: Any, label: str) -> date:
    try:
        return date.fromisoformat(str(value))
    except (TypeError, ValueError):
        raise InvoiceValidationError(f"{label} must be a valid date.") from None


@dataclass(frozen=True)
class InvoiceRequest:
    invoice_type: str
    purpose: str
    date_from: date
    date_to: date
    min_amount: Decimal
    max_amount: Decimal
    currency: str
    tax_type: str
    tax_rate: Decimal
    line_item_mode: str
    max_line_items: int
    company_mode: str
    seller_name: str = ""
    seller_address: str = ""
    seller_tax_id: str = ""
    buyer_name: str = ""
    buyer_address: str = ""
    buyer_tax_id: str = ""
    count: int = 1
    generation_mode: str = "Single"
    output_destination: str = "PDF in Invoice Library"

    @classmethod
    def from_form(
        cls,
        data: Mapping[str, Any],
        max_batch_size: int = 50,
        max_line_items: int = 20,
        max_amount: Decimal = Decimal("100000000"),
    ) -> "InvoiceRequest":
        def choice(key: str, allowed: tuple[str, ...], default: str) -> str:
            value = str(data.get(key) or default).strip()
            if value not in allowed:
                raise InvoiceValidationError(f"Choose a valid {key.replace('_', ' ')}.")
            return value

        try:
            count = int(str(data.get("count") or "1"))
            max_items = int(str(data.get("max_line_items") or "3"))
        except ValueError:
            raise InvoiceValidationError("Invoice count and line item count must be whole numbers.") from None
        if count < 1 or count > max_batch_size:
            raise InvoiceValidationError(f"Invoice count must be between 1 and {max_batch_size}.")
        generation_mode = str(data.get("generation_mode") or ("Single" if count == 1 else "Batch"))
        if generation_mode not in ("Single", "Batch") or (generation_mode == "Single" and count != 1) or (generation_mode == "Batch" and count < 2):
            raise InvoiceValidationError("Choose Single for one invoice or Batch for two or more invoices.")
        output_destination = str(data.get("output_destination") or "PDF in Invoice Library")
        if output_destination != "PDF in Invoice Library":
            raise InvoiceValidationError("Output must be PDF in the Invoice Library.")
        if max_items < 1 or max_items > max_line_items:
            raise InvoiceValidationError(f"Line item count must be between 1 and {max_line_items}.")
        invoice_type = choice("invoice_type", INVOICE_TYPES, "Standard")
        tax_type = choice("tax_type", TAX_TYPES, "None")
        tax_rate = money(data.get("tax_rate") or "0", "Tax rate")
        if tax_rate < 0 or tax_rate > 100:
            raise InvoiceValidationError("Tax rate must be between 0 and 100 percent.")
        if tax_type == "None" and tax_rate != 0:
            raise InvoiceValidationError("Choose a tax type when the tax rate is above zero.")
        if tax_type != "None" and tax_rate == 0:
            raise InvoiceValidationError("Enter a tax rate above zero for the selected tax type.")
        if invoice_type == "GST" and tax_type == "None":
            raise InvoiceValidationError("GST invoices need CGST/SGST or IGST tax.")
        names = {key: str(data.get(key) or "").strip() for key in (
            "seller_name", "seller_address", "seller_tax_id", "buyer_name", "buyer_address", "buyer_tax_id"
        )}
        company_mode = choice("company_mode", ("Generated", "Custom"), "Generated")
        if any(names.values()):
            company_mode = "Custom"
        if any(len(value) > 160 for value in names.values()):
            raise InvoiceValidationError("Company details must be 160 characters or fewer per field.")
        if company_mode == "Custom" and (not names["seller_name"] or not names["buyer_name"]):
            raise InvoiceValidationError("Enter both seller and buyer names for customized invoices.")
        request = cls(
            invoice_type=invoice_type,
            purpose=choice("purpose", PURPOSES, "Testing"),
            date_from=_date(data.get("date_from"), "Start date"),
            date_to=_date(data.get("date_to"), "End date"),
            min_amount=money(data.get("min_amount"), "Minimum amount"),
            max_amount=money(data.get("max_amount"), "Maximum amount"),
            currency=choice("currency", CURRENCIES, "INR"),
            tax_type=tax_type,
            tax_rate=tax_rate,
            line_item_mode=choice("line_item_mode", LINE_ITEM_MODES, "Mixed"),
            max_line_items=max_items,
            company_mode=company_mode,
            count=count,
            generation_mode=generation_mode,
            output_destination=output_destination,
            **names,
        )
        if request.date_from > request.date_to:
            raise InvoiceValidationError("Start date must be on or before end date.")
        if request.min_amount <= 0 or request.max_amount <= 0 or request.min_amount > request.max_amount:
            raise InvoiceValidationError("Enter a positive amount range with minimum at or below maximum.")
        if request.max_amount > max_amount:
            raise InvoiceValidationError(f"Maximum amount must be {max_amount:,.2f} or less.")
        if request.company_mode == "Custom" and request.count > int((request.max_amount - request.min_amount) * 100) + 1:
            raise InvoiceValidationError("Widen the amount range so each invoice in the batch can have a different total.")
        return request

    def to_form(self) -> dict[str, str]:
        return {key: str(value) for key, value in asdict(self).items()}


@dataclass(frozen=True)
class RandomGenerationRequest:
    count: int
    purpose: str

    @classmethod
    def from_form(cls, data: Mapping[str, Any], max_batch_size: int, default_purpose: str) -> "RandomGenerationRequest":
        try:
            number = Decimal(str(data.get("invoice_count") or "").strip())
            if not number.is_finite() or number != number.to_integral_value():
                raise ValueError
            count = int(number)
        except (InvalidOperation, TypeError, ValueError):
            raise InvoiceValidationError("Number of invoices must be a whole number.") from None
        if not 1 <= count <= max_batch_size:
            raise InvoiceValidationError(f"Number of invoices must be between 1 and {max_batch_size}.")
        purpose = str(data.get("purpose") or default_purpose).strip()
        if purpose not in PURPOSES:
            raise InvoiceValidationError("Choose a valid purpose.")
        return cls(count=count, purpose=purpose)


@dataclass(frozen=True)
class Party:
    name: str
    address: str
    tax_id: str = ""


@dataclass(frozen=True)
class LineItem:
    description: str
    quantity: int
    unit_price: Decimal
    amount: Decimal


@dataclass(frozen=True)
class TaxBreakdown:
    tax_type: str
    rate: Decimal
    cgst: Decimal
    sgst: Decimal
    igst: Decimal


@dataclass(frozen=True)
class InvoiceTotals:
    subtotal: Decimal
    discount: Decimal
    taxable_amount: Decimal
    tax: TaxBreakdown
    grand_total: Decimal


@dataclass(frozen=True)
class Invoice:
    invoice_number: str
    invoice_type: str
    purpose: str
    invoice_date: date
    due_date: date
    currency: str
    seller: Party
    buyer: Party
    line_items: tuple[LineItem, ...]
    totals: InvoiceTotals
    batch_id: str
    generated_at: datetime
    generated_by: str

    def to_dict(self) -> dict[str, Any]:
        def encode(value: Any) -> Any:
            if isinstance(value, (date, datetime)):
                return value.isoformat()
            if isinstance(value, Decimal):
                return str(value)
            if isinstance(value, dict):
                return {key: encode(item) for key, item in value.items()}
            if isinstance(value, (list, tuple)):
                return [encode(item) for item in value]
            return value
        return encode(asdict(self))


@dataclass(frozen=True)
class StoredInvoice:
    invoice: Invoice
    item_id: str
    web_url: str


@dataclass(frozen=True)
class FailedInvoice:
    index: int
    stage: str
    message: str


@dataclass(frozen=True)
class BatchResult:
    batch_id: str
    requested_count: int
    generated_count: int
    uploaded_count: int
    failed_items: tuple[FailedInvoice, ...]
    sharepoint_records: tuple[StoredInvoice, ...]

    @property
    def failed_count(self) -> int:
        return self.requested_count - self.uploaded_count

    @property
    def sharepoint_links(self) -> tuple[str, ...]:
        return tuple(record.web_url for record in self.sharepoint_records)
