"""Unit tests for slide-loop tag scanning and tree parsing (pure, no pptx)."""

import pytest

from pptxtpl.exceptions import InvalidTemplateError
from pptxtpl.slide_loops import (
    SLIDE_FOR_RE,
    SLIDE_ENDFOR_RE,
    scan_slide_tags,
    parse_loop_tree,
    expand_loop_tree,
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


def _eval(data):
    """Build an eval_iterable callback over a dict of iterables.

    Expressions of the form "name" look up data[name]; expressions of the
    form "var.attr" look up path_ctx[var][attr] — enough to exercise lazy
    per-iteration evaluation without Jinja.
    """
    def eval_iterable(expr, path_ctx):
        if "." in expr:
            var, attr = expr.split(".", 1)
            return list(path_ctx[var][attr])
        return list(data[expr])
    return eval_iterable


class TestExpandLoopTree:
    def test_single_slide_loop(self):
        roots = parse_loop_tree([[("for", ["x"], "xs", None), ("endfor",)]])
        out = expand_loop_tree(roots[0], _eval({"xs": ["a", "b"]}))
        assert [(i, c["x"]) for i, c in out] == [(0, "a"), (0, "b")]
        assert out[0][1]["loop"] == {
            "index": 1, "index0": 0, "first": True, "last": False, "length": 2,
        }
        assert out[1][1]["loop"]["last"] is True

    def test_multi_slide_group(self):
        roots = parse_loop_tree(
            [[("for", ["x"], "xs", None)], [], [("endfor",)]]
        )
        out = expand_loop_tree(roots[0], _eval({"xs": ["a", "b"]}))
        # Both items produce the full 3-slide group, in order
        assert [i for i, _ in out] == [0, 1, 2, 0, 1, 2]
        assert out[0][1]["x"] == "a"
        assert out[3][1]["x"] == "b"

    def test_nested_cartesian_expansion(self):
        # slides: 0=outer for, 1=inner for, 2=inner endfor, 3=outer endfor
        roots = parse_loop_tree([
            [("for", ["r"], "regions", "rloop")],
            [("for", ["c"], "r.cities", None)],
            [("endfor",)],
            [("endfor",)],
        ])
        regions = [
            {"name": "West", "cities": ["SF", "LA"]},
            {"name": "East", "cities": ["NYC"]},
        ]
        out = expand_loop_tree(roots[0], _eval({"regions": regions}))
        # Per region: slide 0 (outer-only), then per city: slides 1,2,
        # then slide 3 (outer-only)
        assert [i for i, _ in out] == [0, 1, 2, 1, 2, 3, 0, 1, 2, 3]
        # Inner slide contexts carry both region and city
        first_inner = out[1][1]
        assert first_inner["r"]["name"] == "West"
        assert first_inner["c"] == "SF"
        # loop = innermost (city loop)
        assert first_inner["loop"]["length"] == 2
        # Named outer helper reachable from inner slide
        assert first_inner["rloop"]["index"] == 1
        # Outer-only slide has loop = outer helper
        assert out[0][1]["loop"]["length"] == 2
        assert out[0][1]["loop"] is out[0][1]["rloop"]
        # Second region: inner loop has 1 city
        second_region_inner = out[7][1]
        assert second_region_inner["c"] == "NYC"
        assert second_region_inner["rloop"]["index"] == 2
        assert second_region_inner["loop"]["length"] == 1

    def test_empty_inner_prunes_subtree(self):
        roots = parse_loop_tree([
            [("for", ["r"], "regions", None)],
            [("for", ["c"], "r.cities", None)],
            [("endfor",)],
            [("endfor",)],
        ])
        regions = [{"cities": []}, {"cities": ["X"]}]
        out = expand_loop_tree(roots[0], _eval({"regions": regions}))
        # Region 1: only outer slides 0,3; region 2: full expansion
        assert [i for i, _ in out] == [0, 3, 0, 1, 2, 3]

    def test_empty_outer_emits_nothing(self):
        roots = parse_loop_tree(
            [[("for", ["x"], "xs", None), ("endfor",)]]
        )
        assert expand_loop_tree(roots[0], _eval({"xs": []})) == []

    def test_multi_var_unpacking(self):
        roots = parse_loop_tree(
            [[("for", ["k", "v"], "pairs", None), ("endfor",)]]
        )
        out = expand_loop_tree(
            roots[0], _eval({"pairs": [("a", 1), ("b", 2)]})
        )
        assert out[0][1]["k"] == "a"
        assert out[0][1]["v"] == 1
        assert out[1][1]["k"] == "b"

    def test_three_level_nesting(self):
        roots = parse_loop_tree([
            [("for", ["a"], "As", "aloop"),
             ("for", ["b"], "a.Bs", "bloop"),
             ("for", ["c"], "b.Cs", None),
             ("endfor",), ("endfor",), ("endfor",)],
        ])
        As = [{"Bs": [{"Cs": ["x", "y"]}]}]
        out = expand_loop_tree(roots[0], _eval({"As": As}))
        assert len(out) == 2  # two leaf emissions, all on slide 0
        ctx = out[1][1]
        assert ctx["c"] == "y"
        assert ctx["aloop"]["index"] == 1
        assert ctx["bloop"]["index"] == 1
        assert ctx["loop"]["index"] == 2

    def test_contexts_are_independent_copies(self):
        roots = parse_loop_tree(
            [[("for", ["x"], "xs", None), ("endfor",)]]
        )
        out = expand_loop_tree(roots[0], _eval({"xs": ["a", "b"]}))
        out[0][1]["mutated"] = True
        assert "mutated" not in out[1][1]

    def test_two_sibling_children_under_one_outer(self):
        # slides: 0=outer for, 1-2=first inner loop, 3=outer-only,
        # 4-5=second inner loop, 6=outer endfor
        roots = parse_loop_tree([
            [("for", ["r"], "regions", None)],
            [("for", ["c"], "r.cities", None)],
            [("endfor",)],
            [],
            [("for", ["s"], "r.sites", None)],
            [("endfor",)],
            [("endfor",)],
        ])
        regions = [{"cities": ["SF", "LA"], "sites": ["hq"]}]
        out = expand_loop_tree(roots[0], _eval({"regions": regions}))
        # slide 0 (outer), cities: 1,2 twice, slide 3 (outer),
        # sites: 4,5 once, slide 6 (outer)
        assert [i for i, _ in out] == [0, 1, 2, 1, 2, 3, 4, 5, 6]
        # first inner loop ctx has c; second has s; both keep r
        assert out[1][1]["c"] == "SF"
        assert out[3][1]["c"] == "LA"
        assert out[6][1]["s"] == "hq"
        assert out[6][1]["r"]["sites"] == ["hq"]
        # outer-only slides carry the outer loop helper
        assert out[5][1]["loop"]["length"] == 1
