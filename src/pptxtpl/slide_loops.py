"""Slide-level loop parsing and expansion.

Turns {%slide for ... %} / {%slide endfor %} tags scattered across slides
into a bracket-matched loop tree, then expands the tree into a flat,
ordered list of (source_slide_index, context) pairs — the exact clone
sequence.  Pure logic: no pptx objects, so it unit-tests without files.
"""

import re
from dataclasses import dataclass, field

from pptxtpl.exceptions import InvalidTemplateError


# {%slide for VAR[, VAR...] in EXPR [as NAME] %}
SLIDE_FOR_RE = re.compile(
    r"\{%-?\s*slide\s+for\s+(\w+(?:\s*,\s*\w+)*)\s+in\s+(.*?)"
    r"(?:\s+as\s+(\w+))?\s*-?%\}",
    re.DOTALL,
)
SLIDE_ENDFOR_RE = re.compile(r"\{%-?\s*slide\s+endfor\s*-?%\}")


@dataclass
class LoopNode:
    """One {%slide for%}...{%slide endfor%} range in the loop tree."""

    start: int                      # slide index of the for tag
    var_names: list                 # loop variable name(s)
    iterable_expr: str              # Jinja expression for the iterable
    helper_name: object = None      # str from `as NAME`, or None
    end: object = None              # slide index of the endfor tag
    children: list = field(default_factory=list)


def scan_slide_tags(xml_str):
    """Extract slide-loop tag events from one slide's XML, in document order.

    Returns a list of events:
      ("for", var_names_list, iterable_expr, helper_name_or_None)
      ("endfor",)
    """
    events = []
    for m in SLIDE_FOR_RE.finditer(xml_str):
        var_names = [v.strip() for v in m.group(1).split(",")]
        events.append(
            (m.start(), ("for", var_names, m.group(2).strip(), m.group(3)))
        )
    for m in SLIDE_ENDFOR_RE.finditer(xml_str):
        events.append((m.start(), ("endfor",)))
    events.sort(key=lambda e: e[0])
    return [e for _, e in events]


def parse_loop_tree(per_slide_events):
    """Bracket-match per-slide tag events into a list of root LoopNodes.

    ``per_slide_events`` is a list with one entry per slide, each entry a
    list of events as produced by :func:`scan_slide_tags`.

    Raises InvalidTemplateError for unbalanced tags, sibling loops that
    share a slide, or invalid `as` helper names.
    """
    roots = []
    stack = []
    for slide_idx, events in enumerate(per_slide_events):
        for event in events:
            if event[0] == "for":
                _, var_names, expr, helper_name = event
                if helper_name is not None:
                    if helper_name == "loop":
                        raise InvalidTemplateError(
                            "slide loop helper cannot be named 'loop' "
                            f"(slide {slide_idx + 1}); 'loop' always refers "
                            "to the innermost slide loop"
                        )
                    if helper_name in var_names:
                        raise InvalidTemplateError(
                            f"slide loop helper name '{helper_name}' "
                            "collides with the loop variable "
                            f"(slide {slide_idx + 1})"
                        )
                node = LoopNode(
                    start=slide_idx,
                    var_names=var_names,
                    iterable_expr=expr,
                    helper_name=helper_name,
                )
                siblings = stack[-1].children if stack else roots
                if (
                    siblings
                    and siblings[-1].end is not None
                    and siblings[-1].end >= slide_idx
                ):
                    raise InvalidTemplateError(
                        "sibling slide loops cannot share a slide "
                        f"(slide {slide_idx + 1}): a slide is cloned as a "
                        "unit, so it can belong to only one loop"
                    )
                siblings.append(node)
                stack.append(node)
            else:  # endfor
                if not stack:
                    raise InvalidTemplateError(
                        f"{{%slide endfor%}} on slide {slide_idx + 1} has "
                        "no matching {%slide for%}"
                    )
                stack.pop().end = slide_idx
    if stack:
        raise InvalidTemplateError(
            f"{{%slide for%}} on slide {stack[-1].start + 1} has no "
            "matching {%slide endfor%}"
        )
    return roots
