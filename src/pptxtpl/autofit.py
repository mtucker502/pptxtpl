"""Precomputed "shrink text on overflow" autofit for rendered presentations.

PowerPoint stores the autofit shrink factor in the file itself, as
``<a:normAutofit fontScale="..." lnSpcReduction="..."/>``.  It recalculates
that factor only when the text is edited in the UI -- not when the file is
opened.  A generated deck whose text is longer than its template placeholder
therefore overflows the shape until you click into the box and press a key.

This module measures each text frame offline and writes the factor in, so the
deck is laid out correctly the moment it opens.

Measuring requires Pillow, which is an optional dependency::

    pip install "pptxtpl[autofit]"

Usage::

    tpl = PptxTemplate("template.pptx")
    tpl.render(context)
    tpl.save("out.pptx", autofit=True)

or directly against a ``python-pptx`` presentation::

    from pptxtpl.autofit import fit_presentation
    fit_presentation(prs)
"""

from __future__ import annotations

from functools import lru_cache
from typing import NamedTuple

from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Length

from pptxtpl.exceptions import PptxTemplateError

__all__ = [
    "AutofitError",
    "AutofitResult",
    "fit_presentation",
    "fit_shape",
    "fit_slide",
]

# Fallback font size when nothing in the shape, layout, or master declares one.
DEFAULT_SIZE_PT = 18.0

# PowerPoint will not shrink below 25%.
MIN_SCALE = 0.25

# Rendered line height as a multiple of font size, for single-spaced text.
_LINE_HEIGHT = 1.2

# Shave a little off the available width to absorb differences between our
# measuring font and the font PowerPoint will actually use. Applied to width
# rather than height because font metrics affect where lines wrap; stacking a
# second margin on top of the line-height factor shrinks text that in fact fits.
_WIDTH_SAFETY = 0.98

# PowerPoint writes fontScale in 2.5% steps; match it so the value looks native.
_SCALE_STEP = 0.025

# Chrome placeholders are positioned and sized by the master. Rescaling them
# fights the template rather than helping it.
_SKIP_PLACEHOLDERS = frozenset(
    {
        PP_PLACEHOLDER.FOOTER,
        PP_PLACEHOLDER.SLIDE_NUMBER,
        PP_PLACEHOLDER.DATE,
        PP_PLACEHOLDER.HEADER,
    }
)

# Default EMU insets python-pptx applies when a bodyPr omits them.
_DEFAULT_INSET_PT = {"l": 7.2, "r": 7.2, "t": 3.6, "b": 3.6}


class AutofitResult(NamedTuple):
    """One shape that received a precomputed autofit scale."""

    slide_index: int
    shape_name: str
    font_scale: float


class AutofitError(PptxTemplateError):
    """Raised when autofit cannot run (for example, Pillow is missing)."""


@lru_cache(maxsize=256)
def _load_font(size_px: int, font_path: str | None):
    """Return a Pillow font object for measurement at ``size_px``."""
    try:
        from PIL import ImageFont
    except ImportError as exc:  # pragma: no cover - exercised via install extra
        raise AutofitError(
            "Autofit requires Pillow. Install it with: pip install 'pptxtpl[autofit]'"
        ) from exc

    size_px = max(size_px, 1)
    if font_path:
        return ImageFont.truetype(font_path, size_px)
    try:
        # Pillow >= 10.1 bundles a scalable default face.
        return ImageFont.load_default(size=size_px)
    except TypeError as exc:  # pragma: no cover - very old Pillow
        raise AutofitError(
            "Autofit needs Pillow >= 10.1 for a scalable default font, "
            "or an explicit font_path."
        ) from exc


def _inset_pt(text_frame, name: str) -> float:
    """Return one text-frame inset in points, falling back to the PPT default."""
    value = getattr(text_frame, f"margin_{name}")
    if value is None:
        return _DEFAULT_INSET_PT[name[0]]
    return Emu(value).pt


def _lvl1_size_from_list_style(element) -> float | None:
    """Read lvl1pPr/defRPr@sz (in points) from an element's a:lstStyle."""
    lst_style = element.find(qn("a:lstStyle"))
    if lst_style is None:
        return None
    lvl1 = lst_style.find(qn("a:lvl1pPr"))
    if lvl1 is None:
        return None
    def_rpr = lvl1.find(qn("a:defRPr"))
    if def_rpr is None or def_rpr.get("sz") is None:
        return None
    return int(def_rpr.get("sz")) / 100


def _lvl1_size_from_master_styles(master, ph_type) -> float | None:
    """Read the master's titleStyle/bodyStyle/otherStyle level-1 size."""
    tx_styles = master._element.find(qn("p:txStyles"))
    if tx_styles is None:
        return None
    if ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
        style_tag = "p:titleStyle"
    elif ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.SUBTITLE):
        style_tag = "p:bodyStyle"
    else:
        style_tag = "p:otherStyle"
    style = tx_styles.find(qn(style_tag))
    if style is None:
        return None
    lvl1 = style.find(qn("a:lvl1pPr"))
    if lvl1 is None:
        return None
    def_rpr = lvl1.find(qn("a:defRPr"))
    if def_rpr is None or def_rpr.get("sz") is None:
        return None
    return int(def_rpr.get("sz")) / 100


def _inherited_size_pt(shape, fallback: float) -> float:
    """Resolve the effective level-1 font size for a shape.

    Checks, in order: the shape's own list style, the matching layout
    placeholder, the matching master placeholder, and the master text styles.
    """
    own = _lvl1_size_from_list_style(shape.text_frame._txBody)
    if own:
        return own

    if not shape.is_placeholder:
        return fallback

    try:
        idx = shape.placeholder_format.idx
        ph_type = shape.placeholder_format.type
    except (AttributeError, ValueError):
        return fallback

    layout = getattr(shape.part, "slide_layout", None)
    master = getattr(layout, "slide_master", None)

    for source in (layout, master):
        if source is None:
            continue
        for candidate in source.placeholders:
            fmt = candidate.placeholder_format
            if fmt.idx != idx and fmt.type != ph_type:
                continue
            size = _lvl1_size_from_list_style(candidate.text_frame._txBody)
            if size:
                return size

    if master is not None:
        size = _lvl1_size_from_master_styles(master, ph_type)
        if size:
            return size

    return fallback


def _paragraph_size_pt(paragraph, default_pt: float) -> float:
    """Largest explicit run size in a paragraph, else the inherited default."""
    sizes = [run.font.size.pt for run in paragraph.runs if run.font.size is not None]
    if sizes:
        return max(sizes)
    if paragraph.font.size is not None:
        return paragraph.font.size.pt
    return default_pt


@lru_cache(maxsize=65536)
def _word_width(size_px: int, font_path: str | None, word: str) -> float:
    """Width of one word in points, measured once per font size."""
    return _load_font(size_px, font_path).getlength(word)


def _wrapped_line_count(text: str, size_px: int, font_path: str | None, max_width_pt: float) -> int:
    """Count lines produced by wrapping ``text`` at ``max_width_pt``.

    Explicit line breaks are honoured before word wrapping. python-pptx renders
    an ``<a:br/>`` as a vertical tab in ``paragraph.text``; a paragraph built
    only from ``run.text`` would silently lose every hard break it contains,
    which badly undercounts bulleted or address-style blocks.
    """
    if not text.strip():
        return 1

    return sum(
        _wrapped_segment_count(segment, size_px, font_path, max_width_pt)
        for segment in text.replace("\v", "\n").split("\n")
    )


def _wrapped_segment_count(
    text: str, size_px: int, font_path: str | None, max_width_pt: float
) -> int:
    """Count lines produced by greedy word wrapping of a single hard line.

    Each word is measured once and line widths are summed, rather than
    re-measuring the growing line for every word: that was quadratic in the
    words per line, and the binary search repeats the layout ~13 times.
    """
    if not text.strip():
        return 1

    space = _word_width(size_px, font_path, " ")
    lines = 0
    current: float | None = None  # width of the line being filled
    for word in text.split():
        width = _word_width(size_px, font_path, word)
        if current is None:
            current = width
        elif current + space + width <= max_width_pt:
            current += space + width
        else:
            lines += 1
            current = width
    return lines + 1


def _layout(
    text_frame, scale: float, default_pt: float, width_pt: float, font_path: str | None
) -> tuple[float, int]:
    """Lay a text frame out at a font scale.

    Returns:
        ``(total_height_pt, total_line_count)``.
    """
    total = 0.0
    line_count = 0
    for paragraph in text_frame.paragraphs:
        size_pt = _paragraph_size_pt(paragraph, default_pt) * scale
        size_px = max(int(round(size_pt)), 1)
        lines = _wrapped_line_count(paragraph.text, size_px, font_path, width_pt)
        line_count += lines

        line_spacing = paragraph.line_spacing
        if isinstance(line_spacing, Length):  # exact spacing, in points
            line_height = Emu(line_spacing).pt
        elif line_spacing is not None:  # a multiple of single spacing
            line_height = size_pt * line_spacing * _LINE_HEIGHT
        else:
            line_height = size_pt * _LINE_HEIGHT

        total += lines * line_height
        if paragraph.space_before is not None:
            total += Emu(paragraph.space_before).pt
        if paragraph.space_after is not None:
            total += Emu(paragraph.space_after).pt
    return total, line_count


def _fits(
    text_frame,
    scale: float,
    default_pt: float,
    width_pt: float,
    height_pt: float,
    font_path: str | None,
) -> bool:
    """Whether a text frame fits its shape at a given font scale.

    A frame that lays out as a single line always counts as fitting. Designers
    routinely place, say, 30pt text in a 32pt-tall box: the line box is taller
    than the frame, but the glyphs themselves are not, so it looks correct and
    PowerPoint leaves it alone. Shrinking those would be wrong -- the overflow
    worth fixing is text wrapping onto more lines than the shape can show.
    """
    required, lines = _layout(text_frame, scale, default_pt, width_pt, font_path)
    return lines <= 1 or required <= height_pt


def _apply_scale(text_frame, scale: float) -> None:
    """Write a measured scale to a text frame.

    Below 1.0 the frame's autofit setting is replaced with a precomputed
    ``normAutofit``. At 1.0 the existing setting -- ``spAutoFit`` on a
    python-pptx text box, a deliberate ``noAutofit`` -- is left alone, except
    that a stale ``fontScale`` left over from the template's old text is
    cleared so PowerPoint does not keep shrinking text that now fits.
    """
    if scale < 1.0:
        _apply_norm_autofit(text_frame, scale, 0.1 if scale < 0.9 else 0.0)
        return
    existing = text_frame._txBody.bodyPr.find(qn("a:normAutofit"))
    if existing is not None:
        existing.attrib.pop("fontScale", None)
        existing.attrib.pop("lnSpcReduction", None)


def _apply_norm_autofit(text_frame, font_scale: float, lnspc_reduction: float) -> None:
    """Replace a shape's autofit setting with a precomputed normAutofit."""
    body_pr = text_frame._txBody.bodyPr

    for tag in ("a:noAutofit", "a:normAutofit", "a:spAutoFit"):
        existing = body_pr.find(qn(tag))
        if existing is not None:
            body_pr.remove(existing)

    node = body_pr.makeelement(qn("a:normAutofit"), {})
    if font_scale < 1.0:
        node.set("fontScale", str(int(round(font_scale * 100000))))
    if lnspc_reduction > 0:
        node.set("lnSpcReduction", str(int(round(lnspc_reduction * 100000))))

    # Schema order: prstTxWarp, then the autofit choice, then scene3d/sp3d.
    warp = body_pr.find(qn("a:prstTxWarp"))
    position = list(body_pr).index(warp) + 1 if warp is not None else 0
    body_pr.insert(position, node)


def _measure_scale(
    shape,
    default_size_pt: float,
    font_path: str | None,
    min_scale: float,
) -> float | None:
    """Find the font scale at which a shape's text fits, without applying it.

    Returns:
        The scale, or ``None`` if the shape is not a candidate for autofit.
    """
    if not shape.has_text_frame:
        return None
    if shape.is_placeholder and shape.placeholder_format.type in _SKIP_PLACEHOLDERS:
        return None

    text_frame = shape.text_frame
    if not text_frame.text.strip():
        return None

    width_pt = Emu(shape.width).pt - _inset_pt(text_frame, "left") - _inset_pt(text_frame, "right")
    height_pt = Emu(shape.height).pt - _inset_pt(text_frame, "top") - _inset_pt(text_frame, "bottom")
    if width_pt <= 0 or height_pt <= 0:
        return None
    width_pt *= _WIDTH_SAFETY

    default_pt = _inherited_size_pt(shape, default_size_pt)

    if _fits(text_frame, 1.0, default_pt, width_pt, height_pt, font_path):
        return 1.0

    low, high = min_scale, 1.0
    for _ in range(12):
        mid = (low + high) / 2
        if _fits(text_frame, mid, default_pt, width_pt, height_pt, font_path):
            low = mid
        else:
            high = mid

    # Round down to PowerPoint's step so we never land back in overflow.
    return max(int(low / _SCALE_STEP) * _SCALE_STEP, min_scale)


def fit_shape(
    shape,
    *,
    default_size_pt: float = DEFAULT_SIZE_PT,
    font_path: str | None = None,
    min_scale: float = MIN_SCALE,
) -> float | None:
    """Give one shape a precomputed autofit scale.

    Args:
        shape: A shape from a slide. Shapes without text, and footer/slide
            number/date/header placeholders, are skipped.
        default_size_pt: Font size assumed when neither the run, the layout,
            nor the master declares one.
        font_path: Path to a .ttf/.otf used for measurement. Defaults to
            Pillow's bundled face. Supplying the deck's actual font improves
            accuracy.
        min_scale: Lower bound on the shrink factor. PowerPoint uses 0.25.

    Returns:
        The scale applied, or ``None`` if the shape was skipped.
    """
    scale = _measure_scale(shape, default_size_pt, font_path, min_scale)
    if scale is None:
        return None
    _apply_scale(shape.text_frame, scale)
    return scale


def _iter_text_shapes(shapes):
    """Yield text-bearing shapes, descending into groups."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_text_shapes(shape.shapes)
        else:
            yield shape


def _free_bottom_emu(shape, siblings, slide_height_emu: int) -> int:
    """Lowest point a shape could grow to before colliding with another shape.

    Only shapes that start below this one and overlap it horizontally can block
    it; everything else is beside it, not under it.
    """
    try:
        top, height = shape.top, shape.height
        left, width = shape.left, shape.width
    except (AttributeError, TypeError):
        return slide_height_emu
    if None in (top, height, left, width):
        return slide_height_emu

    bottom, right = top + height, left + width
    limit = slide_height_emu
    for other in siblings:
        if other is shape:
            continue
        try:
            o_top, o_left, o_width = other.top, other.left, other.width
        except (AttributeError, TypeError):
            continue
        if None in (o_top, o_left, o_width):
            continue
        if o_top < bottom:
            continue
        if o_left >= right or o_left + o_width <= left:
            continue
        limit = min(limit, o_top)
    return limit


def _grow_to_fit(shape, siblings, slide_height_emu: int, options: dict) -> None:
    """Expand a shape downward into free space so its text can stay larger.

    A cramped box, not the slide, is usually what forces a heavy shrink: a long
    title in a 33pt frame has to drop to ~50% even when there is empty space
    beneath it. Growing the box first spends that space on legibility.

    The shape is only kept at its new height if the extra room actually buys a
    larger font; otherwise it is restored.
    """
    if not shape.has_text_frame:
        return
    # Growing a middle- or bottom-anchored frame would move its text.
    if shape.text_frame.vertical_anchor not in (None, MSO_ANCHOR.TOP):
        return

    before = _measure_scale(shape, **options)
    if before is None or before >= 1.0:
        return

    original = shape.height
    limit = _free_bottom_emu(shape, siblings, slide_height_emu)
    grown = limit - shape.top
    if grown <= original:
        return

    # A placeholder may inherit its geometry from the layout, with no <a:xfrm>
    # of its own. Writing one dimension creates a partial xfrm and stops that
    # inheritance for the rest, leaving width unresolvable -- so pin all four.
    shape.left, shape.top, shape.width = shape.left, shape.top, shape.width
    shape.height = grown
    if (_measure_scale(shape, **options) or 0) <= before:
        shape.height = original


def fit_slide(
    slide,
    *,
    uniform: bool = False,
    grow: bool = False,
    slide_height_emu: int | None = None,
    **kwargs,
) -> list[tuple[str, float]]:
    """Apply :func:`fit_shape` to every text shape on a slide.

    Args:
        slide: The slide to fit.
        uniform: Scale every shape on the slide by the same factor -- the
            smallest any one shape needs -- so text keeps its relative sizes.
            Without it each shape is scaled independently, which can leave a
            cramped title much smaller than the body beneath it.
        grow: Before shrinking, let a shape expand downward into empty space,
            so a tight box is not what forces the shrink. Changes shape
            geometry; only top-anchored frames are touched.
        slide_height_emu: Slide height, needed for ``grow``. Taken from the
            presentation by :func:`fit_presentation`.
        **kwargs: Forwarded to :func:`fit_shape`.

    Returns:
        ``(shape_name, font_scale)`` for shapes that were shrunk.
    """
    options = {
        "default_size_pt": kwargs.get("default_size_pt", DEFAULT_SIZE_PT),
        "font_path": kwargs.get("font_path"),
        "min_scale": kwargs.get("min_scale", MIN_SCALE),
    }
    shapes = list(_iter_text_shapes(slide.shapes))

    if grow:
        if slide_height_emu is None:
            raise AutofitError("grow requires slide_height_emu")
        obstacles = list(slide.shapes)
        for shape in shapes:
            _grow_to_fit(shape, obstacles, slide_height_emu, options)

    if not uniform:
        applied = []
        for shape in shapes:
            scale = fit_shape(shape, **kwargs)
            if scale is not None and scale < 1.0:
                applied.append((shape.name, scale))
        return applied

    measured = [
        (shape, scale)
        for shape in shapes
        if (scale := _measure_scale(shape, **options)) is not None
    ]
    if not measured:
        return []

    slide_scale = min(scale for _, scale in measured)
    applied = []
    for shape, _ in measured:
        _apply_scale(shape.text_frame, slide_scale)
        if slide_scale < 1.0:
            applied.append((shape.name, slide_scale))
    return applied


def fit_presentation(prs, **kwargs) -> list[AutofitResult]:
    """Apply :func:`fit_shape` to every text shape in a presentation.

    Args:
        prs: A ``python-pptx`` ``Presentation``.
        **kwargs: Forwarded to :func:`fit_slide`, so ``uniform`` selects
            per-slide rather than per-shape scaling.

    Returns:
        One :class:`AutofitResult` per shape that was shrunk.
    """
    results = []
    kwargs.setdefault("slide_height_emu", prs.slide_height)
    for slide_index, slide in enumerate(prs.slides, start=1):
        for shape_name, scale in fit_slide(slide, **kwargs):
            results.append(AutofitResult(slide_index, shape_name, scale))
    return results
