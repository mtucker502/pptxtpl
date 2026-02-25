"""RichText and Listing helpers for pptxtpl.

RichText lets users build styled inline text that renders as proper
PowerPoint XML runs (<a:r>) with formatting attributes.

Listing preserves newlines and paragraph breaks for post-processing.
"""

from xml.sax.saxutils import escape


def _emu(pt: float) -> int:
    """Convert points to EMUs (English Metric Units). 1 pt = 12700 EMU."""
    return int(pt * 12700)


def _color_attr(color: str) -> str:
    """Build a solidFill color element from a hex color string like 'FF0000'."""
    color = color.lstrip("#")
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


class RichText:
    """Builds styled text as PowerPoint XML runs.

    Usage::

        rt = RichText("Hello", bold=True, color="FF0000")
        rt.add(" World", italic=True)
        # Use rt in template context — it renders as XML runs
    """

    def __init__(
        self,
        text: str = "",
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        color: str | None = None,
        size: float | None = None,
        font: str | None = None,
    ):
        self._runs: list[str] = []
        if text:
            self.add(
                text,
                bold=bold,
                italic=italic,
                underline=underline,
                color=color,
                size=size,
                font=font,
            )

    def add(
        self,
        text: str,
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        color: str | None = None,
        size: float | None = None,
        font: str | None = None,
    ) -> "RichText":
        """Add a styled run of text."""
        rpr_attrs: list[str] = []
        rpr_children: list[str] = []

        if bold is not None:
            rpr_attrs.append(f' b="{1 if bold else 0}"')
        if italic is not None:
            rpr_attrs.append(f' i="{1 if italic else 0}"')
        if underline is not None:
            rpr_attrs.append(f' u="{"sng" if underline else "none"}"')
        if size is not None:
            rpr_attrs.append(f' sz="{_emu(size) // 100}"')  # hundredths of a point
        if color is not None:
            rpr_children.append(_color_attr(color))
        if font is not None:
            rpr_children.append(
                f'<a:latin typeface="{escape(font)}"/>'
                f'<a:cs typeface="{escape(font)}"/>'
            )

        rpr_xml = ""
        if rpr_attrs or rpr_children:
            rpr_xml = f'<a:rPr{"".join(rpr_attrs)}>{"".join(rpr_children)}</a:rPr>'

        escaped_text = escape(text)
        run_xml = f"<a:r>{rpr_xml}<a:t>{escaped_text}</a:t></a:r>"
        self._runs.append(run_xml)
        return self

    def __str__(self) -> str:
        return "".join(self._runs)


class Listing:
    """Wraps text for multi-line rendering in templates.

    Newlines (``\\n``) in the text become line breaks (``<a:br/>``) during
    post-processing. The bell character (``\\a``) creates paragraph breaks.
    """

    def __init__(self, text: str = ""):
        self._text = text

    def __str__(self) -> str:
        return escape(self._text)
