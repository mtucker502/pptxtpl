"""XML preprocessing utilities for pptxtpl.

Handles the core challenge: PowerPoint splits text across multiple XML <a:r> run
elements, fragmenting Jinja2 tags like {{ and {% %}. These functions reconstitute
fragments before Jinja2 can process them.
"""

import re
from html import unescape

# PowerPoint XML namespaces
NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}

# Regex for content between <a:t> and </a:t> (run boundaries inside Jinja tags)
_RUN_BOUNDARY = re.compile(r"</a:t>.*?<a:t[^>]*>", re.DOTALL)

# Regex matching Jinja2 tags: {{ ... }}, {% ... %}, {# ... #}
_JINJA_TAG = re.compile(r"(\{\{.*?\}\}|\{%.*?%\}|\{#.*?#\})", re.DOTALL)

# Special prefixes and their enclosing XML elements
_SPECIAL_PREFIXES = {
    "pp": ("a:p", r"<a:p\b[^>]*>.*?</a:p>"),
    "sp": ("p:sp", r"<p:sp\b[^>]*>.*?</p:sp>"),
    "tr": ("a:tr", r"<a:tr\b[^>]*>.*?</a:tr>"),
    "tc": ("a:tc", r"<a:tc\b[^>]*>.*?</a:tc>"),
}


def clean_jinja_delimiters(xml: str) -> str:
    """Rejoin Jinja2 delimiters that PowerPoint split across XML runs.

    PowerPoint may insert XML tags between { and {, or % and }, etc.
    This removes those internal tags to reconstitute the delimiters.

    For example, ``{</a:t></a:r><a:r><a:t>{`` becomes ``{{``.
    """
    # Fix split {{ and }}
    xml = re.sub(
        r"\{(<[^>]*>)*\{",
        "{{",
        xml,
    )
    xml = re.sub(
        r"\}(<[^>]*>)*\}",
        "}}",
        xml,
    )
    # Fix split {% and %}
    xml = re.sub(
        r"\{(<[^>]*>)*%",
        "{%",
        xml,
    )
    xml = re.sub(
        r"%(<[^>]*>)*\}",
        "%}",
        xml,
    )
    # Fix split {# and #}
    xml = re.sub(
        r"\{(<[^>]*>)*#",
        "{#",
        xml,
    )
    xml = re.sub(
        r"#(<[^>]*>)*\}",
        "#}",
        xml,
    )
    return xml


def strip_internal_tags(xml: str) -> str:
    """Remove run boundaries (</a:t>...<a:t>) inside Jinja2 expressions.

    After delimiters are rejoined, Jinja tags may still span multiple runs.
    This collapses them so the entire expression is in a single <a:t>.
    """

    def _strip_boundaries(match: re.Match) -> str:
        tag = match.group(0)
        return _RUN_BOUNDARY.sub("", tag)

    return _JINJA_TAG.sub(_strip_boundaries, xml)


def ensure_space_preservation(xml: str) -> str:
    """Add xml:space="preserve" to <a:t> elements containing Jinja2 tags.

    Without this attribute, PowerPoint may strip leading/trailing whitespace
    from text, breaking Jinja2 expressions that rely on spacing.
    """
    # Find <a:t> elements that contain Jinja tags but lack xml:space="preserve"
    def _add_preserve(match: re.Match) -> str:
        opening_tag = match.group(1)
        content = match.group(2)
        if _JINJA_TAG.search(content) and 'xml:space="preserve"' not in opening_tag:
            # Replace <a:t> or <a:t ...> with version that has xml:space="preserve"
            if opening_tag == "<a:t>":
                return f'<a:t xml:space="preserve">{content}</a:t>'
            else:
                return opening_tag.replace("<a:t ", '<a:t xml:space="preserve" ') + content + "</a:t>"
        return match.group(0)

    return re.sub(r"(<a:t[^>]*>)(.*?)</a:t>", _add_preserve, xml, flags=re.DOTALL)


def elevate_special_tags(xml: str) -> str:
    """Replace enclosing XML elements for special-prefix Jinja tags.

    Tags like {%pp if show %} should operate at the paragraph (<a:p>) level.
    This replaces the entire <a:p>...</a:p> containing the tag with just the
    bare Jinja directive (without the prefix).

    Supported prefixes: pp (paragraph), sp (shape), tr (table row), tc (table cell).
    """
    for prefix, (element_tag, _) in _SPECIAL_PREFIXES.items():
        # Pattern: find the enclosing element that contains a special-prefix tag
        # We need to handle nested elements properly by using a non-greedy approach
        # with the specific element tag
        xml = _elevate_prefix(xml, prefix, element_tag)
    return xml


def _elevate_prefix(xml: str, prefix: str, element_tag: str) -> str:
    """Elevate a single prefix's tags to their enclosing XML element level."""
    # Match Jinja tags with this prefix: {%pp ...%} or {{pp ...}}
    tag_pattern = re.compile(
        r"\{[%{]\s*" + re.escape(prefix) + r"\s+(.*?)\s*[%}]\}", re.DOTALL
    )

    # Find all Jinja tags with this prefix
    while True:
        tag_match = tag_pattern.search(xml)
        if not tag_match:
            break

        tag_start = tag_match.start()
        tag_end = tag_match.end()
        full_tag = tag_match.group(0)
        inner = tag_match.group(1)

        # Determine if it's a {% %} or {{ }} tag
        if full_tag.startswith("{%"):
            bare_tag = "{%" + " " + inner + " " + "%}"
        else:
            bare_tag = "{{" + " " + inner + " " + "}}"

        # Find the enclosing element
        open_tag = f"<{element_tag}"
        close_tag = f"</{element_tag}>"

        # Search backwards from the tag for the opening element
        enclosing_start = _find_enclosing_open(xml, tag_start, element_tag)
        if enclosing_start is None:
            # Can't find enclosing element; leave the tag as-is but strip prefix
            xml = xml[:tag_match.start()] + bare_tag + xml[tag_match.end():]
            continue

        # Search forwards from the tag for the closing element
        enclosing_end = _find_enclosing_close(xml, tag_end, element_tag)
        if enclosing_end is None:
            xml = xml[:tag_match.start()] + bare_tag + xml[tag_match.end():]
            continue

        # Replace the entire enclosing element with the bare Jinja tag
        xml = xml[:enclosing_start] + bare_tag + xml[enclosing_end:]

    return xml


def _find_enclosing_open(xml: str, pos: int, element_tag: str) -> int | None:
    """Find the start of the innermost enclosing element of the given tag."""
    open_pattern = re.compile(rf"<{re.escape(element_tag)}[\s>]")
    close_pattern = re.compile(rf"</{re.escape(element_tag)}>")

    # Search backwards: we need to find the matching open tag
    # Count nesting depth going backwards
    depth = 0
    search_pos = pos
    while search_pos > 0:
        # Find the last open or close tag before search_pos
        last_open = None
        for m in open_pattern.finditer(xml, 0, search_pos):
            last_open = m

        last_close = None
        for m in close_pattern.finditer(xml, 0, search_pos):
            last_close = m

        if last_open is None:
            return None

        open_pos = last_open.start() if last_open else -1
        close_pos = last_close.start() if last_close else -1

        if close_pos > open_pos:
            # There's a close tag between us and the open tag — nested element
            depth += 1
            search_pos = close_pos
        else:
            if depth == 0:
                return open_pos
            depth -= 1
            search_pos = open_pos

    return None


def _find_enclosing_close(xml: str, pos: int, element_tag: str) -> int | None:
    """Find the end of the innermost enclosing element's close tag."""
    open_pattern = re.compile(rf"<{re.escape(element_tag)}[\s>]")
    close_tag_str = f"</{element_tag}>"
    close_pattern = re.compile(re.escape(close_tag_str))

    depth = 0
    search_pos = pos
    while search_pos < len(xml):
        next_open = open_pattern.search(xml, search_pos)
        next_close = close_pattern.search(xml, search_pos)

        if next_close is None:
            return None

        open_pos = next_open.start() if next_open else len(xml) + 1
        close_pos = next_close.start()

        if open_pos < close_pos:
            # There's an open tag before the close — nested element
            depth += 1
            search_pos = open_pos + 1
        else:
            if depth == 0:
                return close_pos + len(close_tag_str)
            depth -= 1
            search_pos = close_pos + len(close_tag_str)

    return None


def clean_entities_in_tags(xml: str) -> str:
    """Unescape HTML entities inside Jinja2 tags.

    PowerPoint may encode < > & as &lt; &gt; &amp; and use smart quotes
    inside text. This restores them within Jinja2 expressions only.
    """

    def _unescape_tag(match: re.Match) -> str:
        tag = match.group(0)
        # Unescape HTML entities
        tag = tag.replace("&lt;", "<")
        tag = tag.replace("&gt;", ">")
        tag = tag.replace("&amp;", "&")
        tag = tag.replace("&apos;", "'")
        tag = tag.replace("&quot;", '"')
        # Fix smart quotes (common in PowerPoint)
        tag = tag.replace("\u201c", '"')  # left double quote
        tag = tag.replace("\u201d", '"')  # right double quote
        tag = tag.replace("\u2018", "'")  # left single quote
        tag = tag.replace("\u2019", "'")  # right single quote
        return tag

    return _JINJA_TAG.sub(_unescape_tag, xml)


def preprocess_xml(xml: str) -> str:
    """Run the full preprocessing pipeline on a slide's XML string.

    Steps (in order):
    1. Clean delimiters — rejoin split {{ }}, {% %}, {# #}
    2. Strip internal tags — remove run boundaries inside Jinja expressions
    3. Ensure space preservation — add xml:space="preserve" to relevant <a:t> elements
    4. Elevate special tags — replace enclosing elements for pp/sp/tr/tc prefixed tags
    5. Clean entities — unescape HTML entities inside Jinja expressions
    """
    xml = clean_jinja_delimiters(xml)
    xml = strip_internal_tags(xml)
    xml = ensure_space_preservation(xml)
    xml = elevate_special_tags(xml)
    xml = clean_entities_in_tags(xml)
    return xml
