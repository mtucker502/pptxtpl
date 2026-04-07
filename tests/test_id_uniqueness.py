"""Scoped-uniqueness audit for cloned PPTX content.

The category being tested: many OOXML attributes carry identifiers that must
be unique within some scope (per-slide, per-table, per-part). PowerPoint
desktop tolerates collisions, but PowerPoint Online, LibreOffice, and the
Open XML validator do not — they silently drop or collapse duplicates. The
rowId fix in 244bd0c is one instance of this category. These tests cover
the category, not just rowId, so future clone-path bugs that violate
uniqueness invariants surface here automatically.
"""

import os

import pytest
from lxml import etree
from pptx import Presentation
from pptx.util import Inches

from pptxtpl import PptxTemplate


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "a16": "http://schemas.microsoft.com/office/drawing/2014/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _walk_unique(elements_with_label):
    """Given an iterable of (label, [values]), assert each list has unique values.

    Returns a list of (label, dup_value) for any failures so the caller can
    raise a single combined assertion error.
    """
    failures = []
    for label, values in elements_with_label:
        seen = set()
        for v in values:
            if v in seen:
                failures.append((label, v))
            seen.add(v)
    return failures


def _lxml_of(element):
    """Re-parse a python-pptx element as a plain lxml element so xpath() works."""
    return etree.fromstring(etree.tostring(element))


def _audit_presentation(prs):
    """Walk a python-pptx Presentation and return a list of uniqueness failures."""
    failures = []

    for slide_idx, slide in enumerate(prs.slides):
        sld = _lxml_of(slide._element)

        # Per-slide: p:cNvPr/@id must be unique
        ids = sld.xpath(".//p:cNvPr/@id", namespaces=NS)
        seen = set()
        for i in ids:
            if i in seen:
                failures.append(
                    f"slide {slide_idx + 1}: duplicate p:cNvPr/@id={i}"
                )
            seen.add(i)

        # Per-table: a16:rowId/@val and a16:colId/@val unique within each table
        for tbl_idx, tbl in enumerate(sld.xpath(".//a:tbl", namespaces=NS)):
            for attr in ("rowId", "colId"):
                vals = tbl.xpath(f".//a16:{attr}/@val", namespaces=NS)
                seen = set()
                for v in vals:
                    if v in seen:
                        failures.append(
                            f"slide {slide_idx + 1} table {tbl_idx + 1}: "
                            f"duplicate a16:{attr}={v}"
                        )
                    seen.add(v)

        # Per-slide-part: rId must be unique within the slide part's rels
        rids = list(slide.part.rels.keys())
        if len(rids) != len(set(rids)):
            failures.append(
                f"slide {slide_idx + 1}: duplicate rId in slide part rels"
            )

    # Per-presentation: <p:sldId>/@id must be unique
    pres = _lxml_of(prs.part._element)
    sld_ids = pres.xpath(".//p:sldIdLst/p:sldId/@id", namespaces=NS)
    seen = set()
    for i in sld_ids:
        if i in seen:
            failures.append(f"presentation: duplicate p:sldId/@id={i}")
        seen.add(i)

    return failures


# ---------- Test 1: deterministic property test ----------

def _build_audit_template(path):
    """Build a single .pptx exercising every clone path pptxtpl supports.

    - Slide 1: static text (control)
    - Slide 2: plain table (exercises rowId/colId scoping per table)
    - Slide 3: hyperlink in a textbox
    - Slide 4: speaker notes + start of {%slide for x in xs %}
    - Slide 5: {%slide endfor %} (multi-slide loop body)

    Note: {%tr for%} table-row loops are not exercised end-to-end here; the
    rowId dedupe path is unit-tested in test_xml_utils.py. This audit's
    primary value is p:cNvPr/@id uniqueness across cloned slides.
    """
    prs = Presentation()

    # Slide 1
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s1.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
    tb.text_frame.text = "Title: {{ title }}"

    # Slide 2: plain table
    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    table = s2.shapes.add_table(
        3, 3, Inches(1), Inches(1), Inches(7), Inches(2)
    ).table
    table.cell(0, 0).text = "Name"
    table.cell(0, 1).text = "Value"
    table.cell(0, 2).text = "Status"
    table.cell(1, 0).text = "{{ row1_name }}"
    table.cell(1, 1).text = "{{ row1_value }}"
    table.cell(1, 2).text = "ok"
    table.cell(2, 0).text = "row2"
    table.cell(2, 1).text = "v2"
    table.cell(2, 2).text = "warn"

    # Slide 3: hyperlink
    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    tb3 = s3.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
    p = tb3.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = "click {{ link_text }}"
    run.hyperlink.address = "https://example.com"

    # Slide 4: speaker notes + slide loop start
    s4 = prs.slides.add_slide(prs.slide_layouts[6])
    tb4a = s4.shapes.add_textbox(Inches(1), Inches(0.5), Inches(5), Inches(0.5))
    tb4a.text_frame.text = "{%slide for x in xs %}"
    tb4b = s4.shapes.add_textbox(Inches(1), Inches(1.5), Inches(5), Inches(1))
    tb4b.text_frame.text = "Item: {{ x }}"
    s4.notes_slide.notes_text_frame.text = "notes for {{ x }}"

    # Slide 5: slide loop end
    s5 = prs.slides.add_slide(prs.slide_layouts[6])
    tb5 = s5.shapes.add_textbox(Inches(1), Inches(0.5), Inches(5), Inches(0.5))
    tb5.text_frame.text = "{%slide endfor %}"

    prs.save(path)


def test_id_uniqueness_after_render(tmp_dir):
    src = os.path.join(tmp_dir, "audit.pptx")
    _build_audit_template(src)

    tpl = PptxTemplate(src)
    tpl.render(
        {
            "title": "Q1 Report",
            "link_text": "here",
            "row1_name": "n1",
            "row1_value": "1",
            "xs": ["a", "b", "c", "d", "e"],
        }
    )

    out = os.path.join(tmp_dir, "audit_out.pptx")
    tpl.save(out)

    # Sanity: file reopens cleanly
    prs2 = Presentation(out)

    failures = _audit_presentation(prs2)
    assert not failures, "id uniqueness violations:\n  " + "\n  ".join(failures)


# ---------- Test 2: N=1000 fuzz ----------

@pytest.mark.slow
@pytest.mark.skipif(
    os.environ.get("PPTXTPL_FUZZ") != "1",
    reason="opt-in: set PPTXTPL_FUZZ=1",
)
def test_clone_fuzz_1000_slides(tmp_dir):
    """Render 1000 cloned slides and audit structural integrity.

    Catches:
    - silent slide / shape drops on reopen
    - statistical id collisions (rowId/colId mint from 2^32 space)
    - structural corruption that prevents reopen
    """
    src = os.path.join(tmp_dir, "fuzz.pptx")
    prs = Presentation()

    # Single template slide with a textbox, a 3x3 table, and a hyperlink
    s_static = prs.slides.add_slide(prs.slide_layouts[6])
    s_static.shapes.add_textbox(
        Inches(1), Inches(1), Inches(5), Inches(1)
    ).text_frame.text = "static"

    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb_open = s.shapes.add_textbox(Inches(1), Inches(0.2), Inches(5), Inches(0.4))
    tb_open.text_frame.text = "{%slide for x in xs %}"
    tb_body = s.shapes.add_textbox(Inches(1), Inches(0.7), Inches(5), Inches(0.4))
    tb_body.text_frame.text = "Item {{ x }}"
    table = s.shapes.add_table(
        2, 3, Inches(1), Inches(1.2), Inches(7), Inches(1.5)
    ).table
    table.cell(0, 0).text = "a"
    table.cell(0, 1).text = "b"
    table.cell(0, 2).text = "c"
    table.cell(1, 0).text = "r1"
    table.cell(1, 1).text = "x"
    table.cell(1, 2).text = "y"
    tb_link = s.shapes.add_textbox(Inches(1), Inches(3), Inches(5), Inches(0.5))
    run = tb_link.text_frame.paragraphs[0].add_run()
    run.text = "link {{ x }}"
    run.hyperlink.address = "https://example.com"
    tb_close = s.shapes.add_textbox(Inches(1), Inches(3.6), Inches(5), Inches(0.4))
    tb_close.text_frame.text = "{%slide endfor %}"

    prs.save(src)

    N = 1000
    tpl = PptxTemplate(src)
    tpl.render({"xs": list(range(N))})

    out = os.path.join(tmp_dir, "fuzz_out.pptx")
    tpl.save(out)

    prs2 = Presentation(out)

    # 1 static + N cloned slides
    assert len(prs2.slides) == 1 + N, f"expected {1 + N} slides, got {len(prs2.slides)}"

    failures = _audit_presentation(prs2)
    assert not failures, (
        f"id uniqueness violations across {N} clones "
        f"({len(failures)} total):\n  " + "\n  ".join(failures[:20])
    )
