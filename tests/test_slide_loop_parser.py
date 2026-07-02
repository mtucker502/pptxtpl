"""Unit tests for slide-loop tag scanning and tree parsing (pure, no pptx)."""

import pytest

from pptxtpl.exceptions import InvalidTemplateError
from pptxtpl.slide_loops import (
    SLIDE_FOR_RE,
    SLIDE_ENDFOR_RE,
    scan_slide_tags,
    parse_loop_tree,
)


class TestForRegex:
    def test_basic_for(self):
        m = SLIDE_FOR_RE.search("{%slide for item in items %}")
        assert m.group(1) == "item"
        assert m.group(2).strip() == "items"
        assert m.group(3) is None

    def test_for_with_as(self):
        m = SLIDE_FOR_RE.search("{%slide for jsa in jsas as jsaloop %}")
        assert m.group(1) == "jsa"
        assert m.group(2).strip() == "jsas"
        assert m.group(3) == "jsaloop"

    def test_multi_var_with_as(self):
        m = SLIDE_FOR_RE.search("{%slide for k, v in pairs as kvloop %}")
        assert m.group(1) == "k, v"
        assert m.group(2).strip() == "pairs"
        assert m.group(3) == "kvloop"

    def test_expr_with_filter_no_as(self):
        m = SLIDE_FOR_RE.search("{%slide for x in items | sort %}")
        assert m.group(2).strip() == "items | sort"
        assert m.group(3) is None

    def test_whitespace_control_dashes(self):
        m = SLIDE_FOR_RE.search("{%-slide for x in xs as xl -%}")
        assert m.group(1) == "x"
        assert m.group(3) == "xl"


class TestScanSlideTags:
    def test_orders_tags_by_position(self):
        xml = (
            "<p>{%slide for a in as_ %}</p>"
            "<p>{%slide for b in a.bs as bloop %}</p>"
            "<p>{%slide endfor %}</p>"
            "<p>{%slide endfor %}</p>"
        )
        events = scan_slide_tags(xml)
        assert events == [
            ("for", ["a"], "as_", None),
            ("for", ["b"], "a.bs", "bloop"),
            ("endfor",),
            ("endfor",),
        ]

    def test_no_tags(self):
        assert scan_slide_tags("<p>{{ hello }}</p>") == []


class TestParseLoopTree:
    def test_single_loop_one_slide(self):
        events = [[("for", ["x"], "xs", None), ("endfor",)]]
        roots = parse_loop_tree(events)
        assert len(roots) == 1
        assert roots[0].start == 0
        assert roots[0].end == 0
        assert roots[0].children == []

    def test_multi_slide_loop(self):
        events = [[("for", ["x"], "xs", None)], [], [("endfor",)]]
        roots = parse_loop_tree(events)
        assert roots[0].start == 0
        assert roots[0].end == 2

    def test_nested_loops(self):
        # slide 0: outer for; slide 1: inner for; slide 2: inner endfor;
        # slide 3: outer endfor
        events = [
            [("for", ["r"], "regions", "rloop")],
            [("for", ["c"], "r.cities", None)],
            [("endfor",)],
            [("endfor",)],
        ]
        roots = parse_loop_tree(events)
        assert len(roots) == 1
        outer = roots[0]
        assert (outer.start, outer.end) == (0, 3)
        assert outer.helper_name == "rloop"
        assert len(outer.children) == 1
        inner = outer.children[0]
        assert (inner.start, inner.end) == (1, 2)
        assert inner.iterable_expr == "r.cities"

    def test_inner_single_slide_loop_shares_outer_start(self):
        # outer for + inner for/endfor all on slide 0, outer endfor slide 1
        events = [
            [("for", ["r"], "regions", None),
             ("for", ["c"], "r.cities", None),
             ("endfor",)],
            [("endfor",)],
        ]
        roots = parse_loop_tree(events)
        outer = roots[0]
        assert (outer.start, outer.end) == (0, 1)
        assert (outer.children[0].start, outer.children[0].end) == (0, 0)

    def test_sequential_roots(self):
        events = [
            [("for", ["a"], "as_", None), ("endfor",)],
            [],
            [("for", ["b"], "bs", None), ("endfor",)],
        ]
        roots = parse_loop_tree(events)
        assert len(roots) == 2
        assert roots[0].end == 0
        assert roots[1].start == 2

    def test_endfor_without_for_raises(self):
        with pytest.raises(InvalidTemplateError):
            parse_loop_tree([[("endfor",)]])

    def test_unclosed_for_raises(self):
        with pytest.raises(InvalidTemplateError):
            parse_loop_tree([[("for", ["x"], "xs", None)]])

    def test_sibling_loops_sharing_slide_raises(self):
        # loop1 endfor and loop2 for on the same slide
        events = [
            [("for", ["a"], "as_", None)],
            [("endfor",), ("for", ["b"], "bs", None)],
            [("endfor",)],
        ]
        with pytest.raises(InvalidTemplateError):
            parse_loop_tree(events)

    def test_helper_named_loop_raises(self):
        with pytest.raises(InvalidTemplateError):
            parse_loop_tree([[("for", ["x"], "xs", "loop"), ("endfor",)]])

    def test_helper_colliding_with_var_raises(self):
        with pytest.raises(InvalidTemplateError):
            parse_loop_tree([[("for", ["x"], "xs", "x"), ("endfor",)]])
