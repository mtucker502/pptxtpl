"""PptxTemplate — core render engine for pptxtpl.

Provides the main API for loading a .pptx template, rendering it with a
Jinja2 context dict, and saving the result.
"""

import re
from xml.sax.saxutils import escape

from lxml import etree
from jinja2 import Environment, BaseLoader, TemplateSyntaxError, meta
from pptx import Presentation

from pptxtpl.xml_utils import preprocess_xml
from pptxtpl.richtext import RichText, Listing
from pptxtpl.exceptions import TemplateRenderError, InvalidTemplateError


# Regex for Jinja tags used to discover undeclared variables
_JINJA_TAG_RE = re.compile(r"(\{\{.*?\}\}|\{%.*?%\}|\{#.*?#\})", re.DOTALL)


class PptxTemplate:
    """Load a PowerPoint template, render with Jinja2, and save.

    Usage::

        tpl = PptxTemplate("template.pptx")
        tpl.render({"name": "World"})
        tpl.save("output.pptx")
    """

    def __init__(self, template_path: str):
        self._template_path = template_path
        try:
            self._prs = Presentation(template_path)
        except Exception as exc:
            raise InvalidTemplateError(f"Cannot load template: {exc}") from exc

    @property
    def slides(self):
        """Access the presentation's slides."""
        return self._prs.slides

    def render(self, context: dict | None = None, jinja_env: Environment | None = None) -> None:
        """Render all slides with the given context dict.

        Args:
            context: Template variables dict. RichText and Listing values are
                     automatically converted to their XML/text representations.
            jinja_env: Optional custom Jinja2 Environment. If not provided,
                       a default environment is created.
        """
        if context is None:
            context = {}

        if jinja_env is None:
            jinja_env = Environment(loader=BaseLoader(), autoescape=False)
            jinja_env.globals.update({"RichText": RichText, "Listing": Listing})

        # Convert context values for Jinja2:
        # - RichText/Listing → their XML string representation (already escaped)
        # - Plain strings → XML-escaped to prevent invalid XML after rendering
        render_context = {}
        for key, value in context.items():
            if isinstance(value, (RichText, Listing)):
                render_context[key] = str(value)
            elif isinstance(value, str):
                render_context[key] = escape(value)
            else:
                render_context[key] = value

        for slide in self._prs.slides:
            self._render_slide(slide, render_context, jinja_env)

    def _render_slide(self, slide, context: dict, jinja_env: Environment) -> None:
        """Render a single slide's XML through the Jinja2 pipeline."""
        # Get the slide's XML element
        slide_element = slide._element

        # Serialize to XML string
        xml_str = etree.tostring(slide_element, encoding="unicode")

        # Preprocess: fix fragmented delimiters, strip internal tags, etc.
        xml_str = preprocess_xml(xml_str)

        # Check if there are any Jinja tags after preprocessing
        if not _JINJA_TAG_RE.search(xml_str):
            return  # No templates on this slide

        # Render with Jinja2
        try:
            template = jinja_env.from_string(xml_str)
            rendered_xml = template.render(context)
        except TemplateSyntaxError as exc:
            raise TemplateRenderError(
                f"Jinja2 syntax error on slide: {exc}"
            ) from exc
        except Exception as exc:
            raise TemplateRenderError(
                f"Rendering failed on slide: {exc}"
            ) from exc

        # Post-process: convert \n to line breaks, \a to paragraph breaks
        rendered_xml = self._post_process(rendered_xml)

        # Parse the rendered XML back into an element tree
        try:
            new_element = etree.fromstring(rendered_xml.encode("utf-8"))
        except etree.XMLSyntaxError as exc:
            raise TemplateRenderError(
                f"Rendered XML is invalid: {exc}"
            ) from exc

        # Replace the slide's element tree
        parent = slide_element.getparent()
        if parent is not None:
            parent.replace(slide_element, new_element)
            slide._element = new_element
        else:
            # Slide is the root element — replace children
            slide_element.clear()
            for attr_name, attr_val in new_element.attrib.items():
                slide_element.set(attr_name, attr_val)
            for child in new_element:
                slide_element.append(child)

    def _post_process(self, xml: str) -> str:
        """Convert escape sequences in rendered text to PowerPoint XML.

        - ``\\n`` inside <a:t> elements becomes ``<a:br/>``
        - ``\\a`` becomes a paragraph break (closes and reopens <a:p>)
        """
        # Handle \n → line break within <a:t> elements
        # We replace \n in text content with </a:t></a:r><a:br/><a:r><a:t>
        xml = self._replace_newlines_in_text(xml)
        return xml

    def _replace_newlines_in_text(self, xml: str) -> str:
        """Replace literal \\n characters inside <a:t> elements with <a:br/> elements."""

        def _replace_in_at(match: re.Match) -> str:
            opening = match.group(1)
            content = match.group(2)
            if "\n" not in content:
                return match.group(0)
            # Split on \n and join with line break XML
            parts = content.split("\n")
            result = opening + parts[0] + "</a:t></a:r><a:br/><a:r>" + ("<a:t>".join(
                p + "</a:t></a:r><a:br/><a:r>" for p in parts[1:-1]
            )) + "<a:t>" + parts[-1] + "</a:t>"
            return result

        return re.sub(r"(<a:t[^>]*>)(.*?)</a:t>", _replace_in_at, xml, flags=re.DOTALL)

    def save(self, output_path: str) -> None:
        """Save the rendered presentation to a file."""
        self._prs.save(output_path)

    def get_undeclared_template_variables(
        self, jinja_env: Environment | None = None
    ) -> set[str]:
        """Find all undeclared variables across all slides.

        Returns a set of variable names that appear in Jinja2 expressions
        but are not defined in the environment's globals.
        """
        if jinja_env is None:
            jinja_env = Environment(loader=BaseLoader(), autoescape=False)

        all_vars: set[str] = set()

        for slide in self._prs.slides:
            xml_str = etree.tostring(slide._element, encoding="unicode")
            xml_str = preprocess_xml(xml_str)

            try:
                ast = jinja_env.parse(xml_str)
                variables = meta.find_undeclared_variables(ast)
                all_vars.update(variables)
            except TemplateSyntaxError:
                pass  # Skip slides with syntax errors

        return all_vars
