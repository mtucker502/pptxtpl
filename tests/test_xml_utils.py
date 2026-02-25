"""Tests for xml_utils module — XML preprocessing functions."""

from pptxtpl.xml_utils import (
    clean_jinja_delimiters,
    strip_internal_tags,
    ensure_space_preservation,
    elevate_special_tags,
    clean_entities_in_tags,
    preprocess_xml,
)


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
