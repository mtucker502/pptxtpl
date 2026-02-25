"""Custom exceptions for pptxtpl."""


class PptxTemplateError(Exception):
    """Base exception for pptxtpl errors."""


class TemplateRenderError(PptxTemplateError):
    """Raised when Jinja2 rendering fails."""


class InvalidTemplateError(PptxTemplateError):
    """Raised when the template .pptx is invalid or cannot be loaded."""
