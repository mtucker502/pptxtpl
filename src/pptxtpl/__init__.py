"""pptxtpl — Jinja2 templating for PowerPoint .pptx files."""

from pptxtpl.template import PptxTemplate
from pptxtpl.richtext import RichText, Listing
from pptxtpl.inline_image import InlineImage
from pptxtpl.autofit import AutofitError, AutofitResult, fit_presentation, fit_shape, fit_slide

__all__ = [
    "PptxTemplate",
    "RichText",
    "Listing",
    "InlineImage",
    "AutofitError",
    "AutofitResult",
    "fit_presentation",
    "fit_shape",
    "fit_slide",
]
