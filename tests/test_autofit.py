"""Tests for precomputed autofit (shrink text on overflow)."""

import pytest
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from pptxtpl.autofit import fit_presentation, fit_shape, fit_slide

pytest.importorskip("PIL", reason="autofit requires Pillow")


def _body_pr(shape):
    return shape.text_frame._txBody.bodyPr


def _norm_autofit(shape):
    return _body_pr(shape).find(qn("a:normAutofit"))


def _textbox(slide, text, *, width_in, height_in, size_pt=18):
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(width_in), Inches(height_in))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = text
    for paragraph in tf.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size_pt)
    return box


def test_overflowing_shape_gets_font_scale():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = _textbox(slide, "word " * 300, width_in=3, height_in=1)

    scale = fit_shape(box)

    assert scale is not None
    assert scale < 1.0
    node = _norm_autofit(box)
    assert node is not None
    assert int(node.get("fontScale")) < 100000


def test_short_text_is_not_shrunk():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = _textbox(slide, "Hi", width_in=8, height_in=3)

    scale = fit_shape(box)

    assert scale == 1.0
    node = _norm_autofit(box)
    assert node is not None
    assert node.get("fontScale") is None


def test_single_line_taller_than_its_box_is_not_shrunk():
    # Designers routinely put 30pt text in a box barely 32pt tall. The line box
    # exceeds the frame but the glyphs do not, so it must be left alone.
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = _textbox(slide, "Customer Presentation", width_in=12, height_in=0.45, size_pt=30)

    scale = fit_shape(box)

    assert scale == 1.0
    assert _norm_autofit(box).get("fontScale") is None


def test_wrapping_text_in_a_short_box_is_shrunk():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = _textbox(slide, "Customer Presentation " * 6, width_in=4, height_in=0.45, size_pt=30)

    scale = fit_shape(box)

    assert scale < 1.0


def test_hard_line_breaks_are_counted():
    # An <a:br/> is not a run. Measuring only run text loses every hard break
    # in a paragraph and badly undercounts bulleted blocks.
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = _textbox(slide, "Short line", width_in=8, height_in=1.5, size_pt=18)
    paragraph = box.text_frame.paragraphs[0]
    for _ in range(20):
        paragraph.add_line_break()
        run = paragraph.add_run()
        run.text = "Short line"
        run.font.size = Pt(18)

    assert fit_shape(box) < 1.0


def test_no_autofit_is_replaced():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = _textbox(slide, "word " * 300, width_in=3, height_in=1)

    body_pr = _body_pr(box)
    body_pr.append(body_pr.makeelement(qn("a:noAutofit"), {}))

    fit_shape(box)

    assert body_pr.find(qn("a:noAutofit")) is None
    assert body_pr.find(qn("a:normAutofit")) is not None


def test_scale_never_goes_below_min():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = _textbox(slide, "word " * 5000, width_in=2, height_in=0.4)

    scale = fit_shape(box, min_scale=0.5)

    assert scale == pytest.approx(0.5)


def test_empty_and_sizeless_shapes_are_skipped():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    empty = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    picture_like = slide.shapes.add_shape(1, Inches(1), Inches(3), Inches(1), Inches(1))

    assert fit_shape(empty) is None
    assert fit_shape(picture_like) is None


def test_footer_placeholders_are_skipped():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    _textbox(slide, "word " * 300, width_in=3, height_in=1)

    results = fit_presentation(prs)

    assert all("Footer" not in name and "Slide Number" not in name for _, name, _ in results)


def test_fit_presentation_reports_each_shrunk_shape():
    prs = Presentation()
    for _ in range(3):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _textbox(slide, "word " * 300, width_in=3, height_in=1)

    results = fit_presentation(prs)

    assert len(results) == 3
    assert [r.slide_index for r in results] == [1, 2, 3]
    assert all(r.font_scale < 1.0 for r in results)


def test_save_with_autofit_writes_scale(tmp_path):
    from pptxtpl import PptxTemplate

    template_path = tmp_path / "template.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _textbox(slide, "{{ body }}", width_in=3, height_in=1)
    prs.save(template_path)

    out_path = tmp_path / "out.pptx"
    tpl = PptxTemplate(str(template_path))
    tpl.render({"body": "word " * 300})
    tpl.save(str(out_path), autofit=True)

    rendered = Presentation(str(out_path))
    node = rendered.slides[0].shapes[0].text_frame._txBody.bodyPr.find(qn("a:normAutofit"))
    assert node is not None
    assert int(node.get("fontScale")) < 100000


def test_save_without_autofit_leaves_shape_untouched(tmp_path):
    from pptxtpl import PptxTemplate

    template_path = tmp_path / "template.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _textbox(slide, "{{ body }}", width_in=3, height_in=1)
    prs.save(template_path)

    out_path = tmp_path / "out.pptx"
    tpl = PptxTemplate(str(template_path))
    tpl.render({"body": "word " * 300})
    tpl.save(str(out_path))

    rendered = Presentation(str(out_path))
    assert rendered.slides[0].shapes[0].text_frame._txBody.bodyPr.find(qn("a:normAutofit")) is None


def test_uniform_scales_every_shape_on_the_slide_alike():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    crowded = _textbox(slide, "word " * 200, width_in=3, height_in=1)
    roomy = _textbox(slide, "word " * 20, width_in=3, height_in=1)

    fit_slide(slide, uniform=True)

    assert _norm_autofit(crowded).get("fontScale") == _norm_autofit(roomy).get("fontScale")


def test_uniform_takes_the_smallest_scale_any_shape_needs():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    crowded = _textbox(slide, "word " * 200, width_in=3, height_in=1)
    _textbox(slide, "word " * 20, width_in=3, height_in=1)

    alone = Presentation()
    alone_slide = alone.slides.add_slide(alone.slide_layouts[6])
    reference = _textbox(alone_slide, "word " * 200, width_in=3, height_in=1)
    fit_shape(reference)

    fit_slide(slide, uniform=True)

    assert _norm_autofit(crowded).get("fontScale") == _norm_autofit(reference).get("fontScale")


def test_grow_expands_a_cramped_shape_into_free_space():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = _textbox(slide, "word " * 40, width_in=6, height_in=0.4)
    original_height = box.height

    without = fit_shape(box)
    fit_slide(slide, grow=True, slide_height_emu=prs.slide_height)

    assert box.height > original_height
    # A missing fontScale means no shrink was needed at all.
    raw = _norm_autofit(box).get("fontScale")
    assert (float(raw) / 100000 if raw else 1.0) > without


def test_grow_stops_at_the_shape_below():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    top = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(6), Inches(0.4))
    top.text_frame.word_wrap = True
    top.text_frame.text = "word " * 40
    for run in top.text_frame.paragraphs[0].runs:
        run.font.size = Pt(18)
    blocker = slide.shapes.add_textbox(Inches(0.5), Inches(2.0), Inches(6), Inches(1))
    blocker.text_frame.text = "below"

    fit_slide(slide, grow=True, slide_height_emu=prs.slide_height)

    assert top.top + top.height <= blocker.top


def test_grow_leaves_a_shape_with_nothing_to_gain_alone():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = _textbox(slide, "Hi", width_in=6, height_in=2)
    original_height = box.height

    fit_slide(slide, grow=True, slide_height_emu=prs.slide_height)

    assert box.height == original_height


def test_grow_requires_a_slide_height():
    from pptxtpl.exceptions import PptxTemplateError

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _textbox(slide, "word " * 50, width_in=3, height_in=0.5)

    with pytest.raises(PptxTemplateError):
        fit_slide(slide, grow=True)
