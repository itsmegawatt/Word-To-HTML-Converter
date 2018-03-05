import unittest
import wordtohtml as wth

class TestWordToHTMLConversion(unittest.TestCase):
    
    def setUp(self):
        self.NAMESPACE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        self.NAMESPACE_GROUP = self.NAMESPACE + "r"
        self.NAMESPACE_BOLD = self.NAMESPACE + "b"
        self.NAMESPACE_ITALICS = self.NAMESPACE + "i"
        self.NAMESPACE_UNDERLINE = self.NAMESPACE + "u"

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

    def test_grab_text(self):
        for paragraph in self.paragraphs:
            groups = paragraph.iter(self.NAMESPACE_GROUP)
            for group in groups:
                text = "".join(t.text for t in group.iter(self.NAMESPACE + "t"))
                self.assertEqual(text, wth.grab_text(group))

    def test_run(self):
        docx_path = r'test.docx'
        #Test1: Save HTML the normal way from paragraphs
        paragraphs = wth.retrieve_paragraphs(docx_path)
        html_lines = wth.convert_to_html_lines_from_paragraphs(paragraphs)
        wth.save_as_html(html_lines, "test1")


        #Test2: Save HTML the faster way directly from Path
        html_lines_2 = wth.convert_to_html_lines_from_path(docx_path)
        wth.save_as_html(html_lines_2, "test2.html")

if __name__ == '__main__':
    unittest.main()