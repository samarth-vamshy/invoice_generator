"""Application errors shown to users only through safe messages."""


class InvoiceValidationError(ValueError):
    """A generation request or invoice violates a business rule."""


class InvoiceGenerationError(Exception):
    """A valid request could not be generated."""


class PdfGenerationError(Exception):
    """Rendering or saving a PDF failed."""


class SharePointError(Exception):
    """The authoritative SharePoint record could not be read or stored."""
