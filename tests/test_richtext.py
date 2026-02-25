"""Tests for RichText and Listing classes."""

from pptxtpl.richtext import RichText, Listing


class TestRichText:
    def test_plain_text(self):
        rt = RichText("Hello")
        xml = str(rt)
        assert "<a:r>" in xml
        assert "<a:t>Hello</a:t>" in xml

    def test_bold(self):
        rt = RichText("Bold", bold=True)
        xml = str(rt)
        assert 'b="1"' in xml

    def test_italic(self):
        rt = RichText("Italic", italic=True)
        xml = str(rt)
        assert 'i="1"' in xml

    def test_underline(self):
        rt = RichText("Underline", underline=True)
        xml = str(rt)
        assert 'u="sng"' in xml

    def test_color(self):
        rt = RichText("Red", color="FF0000")
        xml = str(rt)
        assert "FF0000" in xml
        assert "<a:solidFill>" in xml

    def test_color_with_hash(self):
        rt = RichText("Blue", color="#0000FF")
        xml = str(rt)
        assert "0000FF" in xml

    def test_font(self):
        rt = RichText("Custom", font="Arial")
        xml = str(rt)
        assert 'typeface="Arial"' in xml

    def test_size(self):
        rt = RichText("Big", size=24)
        xml = str(rt)
        assert "sz=" in xml

    def test_add_multiple_runs(self):
        rt = RichText("Hello", bold=True)
        rt.add(" World", italic=True)
        xml = str(rt)
        assert xml.count("<a:r>") == 2
        assert "Hello" in xml
        assert "World" in xml

    def test_escapes_special_chars(self):
        rt = RichText("a < b & c > d")
        xml = str(rt)
        assert "&lt;" in xml
        assert "&amp;" in xml
        assert "&gt;" in xml

    def test_empty_richtext(self):
        rt = RichText()
        assert str(rt) == ""

    def test_chained_add(self):
        rt = RichText()
        rt.add("A").add("B").add("C")
        xml = str(rt)
        assert xml.count("<a:r>") == 3


class TestListing:
    def test_plain_text(self):
        lst = Listing("Line 1\nLine 2")
        result = str(lst)
        assert "Line 1\nLine 2" in result

    def test_escapes_special_chars(self):
        lst = Listing("a < b & c")
        result = str(lst)
        assert "&lt;" in result
        assert "&amp;" in result

    def test_empty_listing(self):
        lst = Listing()
        assert str(lst) == ""
