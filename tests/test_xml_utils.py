"""Tests for xml_utils module — XML preprocessing functions."""

import random
import re

from pptxtpl.xml_utils import (
    clean_jinja_delimiters,
    strip_internal_tags,
    ensure_space_preservation,
    elevate_special_tags,
    clean_entities_in_tags,
    preprocess_xml,
    dedupe_table_ids,
)


def _row(rowid: int) -> str:
    return (
        '<a:tr h="100"><a:tc><a:txBody><a:p/></a:txBody></a:tc>'
        f'<a:extLst><a:ext uri="{{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}}">'
        f'<a16:rowId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="{rowid}"/>'
        '</a:ext></a:extLst></a:tr>'
    )


class TestDedupeTableIds:
    def test_duplicate_rowids_are_regenerated(self):
        xml = "<a:tbl>" + _row(123) + _row(123) + _row(123) + "</a:tbl>"
        out = dedupe_table_ids(xml, _rng=random.Random(0))
        ids = [int(v) for v in re.findall(r'a16:rowId[^v]*val="(\d+)"', out)]
        assert len(ids) == 3
        assert len(set(ids)) == 3
        assert ids[0] == 123  # first occurrence preserved

    def test_unique_rowids_are_left_alone(self):
        xml = "<a:tbl>" + _row(1) + _row(2) + _row(3) + "</a:tbl>"
        out = dedupe_table_ids(xml)
        assert out == xml

    def test_duplicate_colids_are_regenerated(self):
        col = (
            '<a:gridCol w="100"><a:extLst><a:ext uri="x">'
            '<a16:colId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="42"/>'
            '</a:ext></a:extLst></a:gridCol>'
        )
        xml = "<a:tbl><a:tblGrid>" + col + col + "</a:tblGrid></a:tbl>"
        out = dedupe_table_ids(xml, _rng=random.Random(0))
        ids = [int(v) for v in re.findall(r'a16:colId[^v]*val="(\d+)"', out)]
        assert len(set(ids)) == 2

    def test_dedup_is_scoped_per_table(self):
        # Same id in two different tables should NOT be regenerated.
        xml = (
            "<a:tbl>" + _row(7) + "</a:tbl>"
            "<a:tbl>" + _row(7) + "</a:tbl>"
        )
        out = dedupe_table_ids(xml)
        assert out == xml


class TestCleanJinjaDelimiters:
    """Tests for rejoining split Jinja2 delimiters."""

    def test_split_double_braces(self):
        xml = '{</a:t></a:r><a:r><a:t>{ name }</a:t></a:r><a:r><a:t>}'
        result = clean_jinja_delimiters(xml)
        assert "{{" in result
        assert "}}" in result

    def test_split_block_tags(self):
        xml = '{</a:t></a:r><a:r><a:t>% if x %</a:t></a:r><a:r><a:t>}'
        result = clean_jinja_delimiters(xml)
        assert "{%" in result
        assert "%}" in result

    def test_split_comment_tags(self):
        xml = '{</a:t></a:r><a:r><a:t># comment #</a:t></a:r><a:r><a:t>}'
        result = clean_jinja_delimiters(xml)
        assert "{#" in result
        assert "#}" in result

    def test_no_split_passthrough(self):
        xml = '<a:t>{{ name }}</a:t>'
        result = clean_jinja_delimiters(xml)
        assert result == xml

    def test_multiple_splits_in_one_string(self):
        xml = (
            '{</a:t><a:t>{ a }</a:t><a:t>} and '
            '{</a:t><a:t>{ b }</a:t><a:t>}'
        )
        result = clean_jinja_delimiters(xml)
        assert result.count("{{") == 2
        assert result.count("}}") == 2


class TestStripInternalTags:
    def test_remove_run_boundaries_in_expression(self):
        xml = '{{ na</a:t></a:r><a:r><a:rPr/><a:t>me }}'
        result = strip_internal_tags(xml)
        assert result == "{{ name }}"

    def test_leave_non_jinja_content_alone(self):
        xml = '<a:t>Hello</a:t></a:r><a:r><a:t>World</a:t>'
        result = strip_internal_tags(xml)
        assert result == xml

    def test_block_tag_with_internal_runs(self):
        xml = '{% i</a:t></a:r><a:r><a:t>f show %}'
        result = strip_internal_tags(xml)
        assert result == "{% if show %}"


class TestEnsureSpacePreservation:
    def test_adds_preserve_to_jinja_tag(self):
        xml = '<a:t>{{ name }}</a:t>'
        result = ensure_space_preservation(xml)
        assert 'xml:space="preserve"' in result
        assert "{{ name }}" in result

    def test_no_change_for_plain_text(self):
        xml = '<a:t>Hello World</a:t>'
        result = ensure_space_preservation(xml)
        assert 'xml:space="preserve"' not in result

    def test_no_duplicate_preserve(self):
        xml = '<a:t xml:space="preserve">{{ name }}</a:t>'
        result = ensure_space_preservation(xml)
        assert result.count('xml:space="preserve"') == 1


class TestElevateSpecialTags:
    def test_pp_prefix_elevates_to_paragraph(self):
        xml = '<a:p><a:r><a:t>{%pp if show %}</a:t></a:r></a:p>'
        result = elevate_special_tags(xml)
        assert "<a:p>" not in result
        assert "{% if show %}" in result
        assert "pp" not in result.replace("{% if show %}", "")

    def test_tr_prefix_elevates_to_table_row(self):
        xml = '<a:tr><a:tc><a:txBody><a:p><a:r><a:t>{%tr for r in rows %}</a:t></a:r></a:p></a:txBody></a:tc></a:tr>'
        result = elevate_special_tags(xml)
        assert "<a:tr>" not in result
        assert "{% for r in rows %}" in result

    def test_no_prefix_not_elevated(self):
        xml = '<a:p><a:r><a:t>{% if show %}</a:t></a:r></a:p>'
        result = elevate_special_tags(xml)
        assert "<a:p>" in result  # paragraph preserved
        assert "{% if show %}" in result


class TestCleanEntitiesInTags:
    def test_unescape_lt_gt(self):
        xml = '{{ x &lt; 10 }}'
        result = clean_entities_in_tags(xml)
        assert result == "{{ x < 10 }}"

    def test_unescape_amp(self):
        xml = '{{ x &amp; y }}'
        result = clean_entities_in_tags(xml)
        assert result == "{{ x & y }}"

    def test_unescape_quotes(self):
        xml = "{{ x == &quot;hello&quot; }}"
        result = clean_entities_in_tags(xml)
        assert result == '{{ x == "hello" }}'

    def test_smart_quotes(self):
        xml = "{{ x == \u201chello\u201d }}"
        result = clean_entities_in_tags(xml)
        assert result == '{{ x == "hello" }}'

    def test_no_change_outside_tags(self):
        xml = '<a:t>5 &lt; 10</a:t> {{ name }}'
        result = clean_entities_in_tags(xml)
        # Entities outside Jinja tags should be preserved
        assert "&lt;" in result
        assert "{{ name }}" in result


class TestPreprocessXml:
    def test_full_pipeline(self):
        # Fragmented tag with entities
        xml = '<a:t>{</a:t></a:r><a:r><a:t>{ na</a:t></a:r><a:r><a:t>me &amp; title }</a:t></a:r><a:r><a:t>}</a:t>'
        result = preprocess_xml(xml)
        assert "{{ name & title }}" in result
