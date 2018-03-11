import unittest
import wordtohtml as wth

class TestWordToHTMLConversion(unittest.TestCase):

    def setUp(self):
        self.NAMESPACE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        self.NAMESPACE_GROUP = self.NAMESPACE + "r"
        self.NAMESPACE_BOLD = self.NAMESPACE + "b"
        self.NAMESPACE_ITALICS = self.NAMESPACE + "i"
        self.NAMESPACE_UNDERLINE = self.NAMESPACE + "u"
        self.NAMESPACE_LINEBREAK = self.NAMESPACE + "br"
        self.NAMESPACE_STRIKE = self.NAMESPACE + "strike"

        self.DOCX_PATH = r'test.docx'
        self.paragraphs = wth.retrieve_paragraphs(self.DOCX_PATH)
        self.text = "Hello"

        # convert_to_html(DOCX_PATH)

    def test_paragraph_text(self):
        self.assertEqual(wth.paragraph_text(self.text), "<p>Hello</p>")

    def test_bold_text(self):
        self.assertEqual(wth.bold_text(self.text), "<b>Hello</b>")

    def test_italicize_text(self):
        self.assertEqual(wth.italicize_text(self.text), "<i>Hello</i>")

    def test_underline_text(self):
        self.assertEqual(wth.underline_text(self.text), "<span style='text-decoration: underline;'>Hello</span>")

    def test_strike_text(self):
        self.assertEqual(wth.strike_text(self.text), "<del>Hello</del>")

    def test_change_symbols_to_character_codes(self):
        self.assertEqual(wth.change_symbols_to_character_codes("&&&"), "&amp;&amp;&amp;")
        self.assertEqual(wth.change_symbols_to_character_codes("‘‘‘"), "&apos;&apos;&apos;")
        self.assertEqual(wth.change_symbols_to_character_codes("’’’"), "&apos;&apos;&apos;")
        self.assertEqual(wth.change_symbols_to_character_codes(":::"), "&colon;&colon;&colon;")
        self.assertEqual(wth.change_symbols_to_character_codes("–––"), "&ndash;&ndash;&ndash;")
        self.assertEqual(wth.change_symbols_to_character_codes("“““"), "&quot;&quot;&quot;")
        self.assertEqual(wth.change_symbols_to_character_codes("”””"), "&quot;&quot;&quot;")
        self.assertEqual(wth.change_symbols_to_character_codes("<<<"), "&lt;&lt;&lt;")
        self.assertEqual(wth.change_symbols_to_character_codes(">>>"), "&gt;&gt;&gt;")
        self.assertEqual(wth.change_symbols_to_character_codes("………"), ".........")
        self.assertEqual(wth.change_symbols_to_character_codes(""), "")
        self.assertEqual(wth.change_symbols_to_character_codes(""), "")

    def test_is_bold(self):
        for paragraph in self.paragraphs:
            groups = paragraph.iter(self.NAMESPACE_GROUP)
            for group in groups:
                bolds = group.iter(self.NAMESPACE_BOLD)
                is_bold = False
                for bold in bolds:
                    is_bold = True
                self.assertEqual(is_bold, wth.is_bold(group))

    def test_is_italics(self):
        for paragraph in self.paragraphs:
            groups = paragraph.iter(self.NAMESPACE_GROUP)
            for group in groups:
                italics = group.iter(self.NAMESPACE_ITALICS)
                is_italics = False
                for italic in italics:
                    is_italics = True
                self.assertEqual(is_italics, wth.is_italics(group))

    def test_is_underline(self):
        for paragraph in self.paragraphs:
            groups = paragraph.iter(self.NAMESPACE_GROUP)
            for group in groups:
                underlines = group.iter(self.NAMESPACE_UNDERLINE)
                is_underline = False
                for underline in underlines:
                    is_underline = True
                self.assertEqual(is_underline, wth.is_underline(group))

    def test_is_strike(self):
        docx_path = r'test3.docx'
        paragraphs = wth.retrieve_paragraphs(docx_path)
        for paragraph in paragraphs:
            groups = paragraph.iter(self.NAMESPACE_GROUP)
            for group in groups:
                strikes = group.iter(self.NAMESPACE_STRIKE)
                is_strike = False
                for strike in strikes:
                    is_strike = True
                self.assertEqual(is_strike, wth.is_strike(group))

    def test_needs_linebreak_before(self):
        docx_path = r'test2.docx'
        paragraphs = wth.retrieve_paragraphs(docx_path)
        for paragraph in paragraphs:
            groups = paragraph.iter(self.NAMESPACE_GROUP)
            for group in groups:
                linebreaks = group.iter(self.NAMESPACE_LINEBREAK)
                needs_linebreak_before = False
                for linebreak in linebreaks:
                    needs_linebreak_before = True
                self.assertEqual(needs_linebreak_before, wth.needs_linebreak_before(group))

    def test_is_blank(self):
        text1 = ''
        text2 = '\n'
        text3 = "not blank"
        self.assertTrue(wth.is_blank(text1))
        self.assertTrue(wth.is_blank(text2))
        self.assertFalse(wth.is_blank(text3))

    def test_grab_text(self):
        for paragraph in self.paragraphs:
            groups = paragraph.iter(self.NAMESPACE_GROUP)
            for group in groups:
                text = "".join(t.text for t in group.iter(self.NAMESPACE + "t"))
                self.assertEqual(text, wth.grab_text(group))

    def convert_to_html_lines_from_paragraphs(self, paragraphs):
        result = []
        for paragraph in paragraphs:
            final_text = ""
            groups = paragraph.iter(self.NAMESPACE_GROUP)
            for group in groups:
                needs_bold = wth.is_bold(group)
                needs_italics = wth.is_italics(group)
                needs_underline = wth.is_underline(group)
                needs_strike = wth.is_strike(group)

                text = wth.grab_text(group)
                text = wth.change_symbols_to_character_codes(text)
                if needs_bold:
                    text = wth.bold_text(text)
                if needs_italics:
                    text = wth.italicize_text(text)
                if needs_underline:
                    text = wth.underline_text(text)
                if needs_strike:
                    text = wth.strike_text(text)

                if wth.needs_linebreak_before(group):
                    final_text += "<br />"
                final_text += text
                #print(final_text)
            if not wth.is_blank(final_text):
                result.append(wth.paragraph_text(final_text))
        return result

    def test_run1(self):
        # Test1: Basic test of the formatting and html output
        # Also tests bold, italics, underline, and symbols
        # Saves HTML the normal way from paragraphs
        docx_path = r'test.docx'
        paragraphs = wth.retrieve_paragraphs(docx_path)
        correct_html_lines = self.convert_to_html_lines_from_paragraphs(paragraphs)
        paragraphs_again = wth.retrieve_paragraphs(docx_path)
        html_lines = wth.convert_to_html_lines_from_paragraphs(paragraphs_again)
        self.assertEqual(correct_html_lines, html_lines)
        wth.save_as_html(html_lines, "test1")

    def test_run2(self):
        # Test2: Tests for linebreaks
        # Saves HTML the faster way directly from Path
        docx_path = r'test2.docx'
        paragraphs = wth.retrieve_paragraphs(docx_path)
        correct_html_lines = self.convert_to_html_lines_from_paragraphs(paragraphs)
        html_lines = wth.convert_to_html_lines_from_path(docx_path)
        self.assertEqual(correct_html_lines, html_lines)
        wth.save_as_html(html_lines, "test2.html")

    def test_run3(self):
        # Test3: Tests for strikethroughs
        docx_path = r'test3.docx'
        paragraphs = wth.retrieve_paragraphs(docx_path)
        correct_html_lines = self.convert_to_html_lines_from_paragraphs(paragraphs)
        paragraphs_again = wth.retrieve_paragraphs(docx_path)
        html_lines = wth.convert_to_html_lines_from_paragraphs(paragraphs_again)
        self.assertEqual(correct_html_lines, html_lines)
        wth.save_as_html(html_lines, "test3.html")

if __name__ == '__main__':
    unittest.main()