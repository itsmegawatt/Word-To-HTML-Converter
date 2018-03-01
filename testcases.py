import unittest
import wordtohtml as wth

class TestWordToHTMLConversion(unittest.TestCase):
    
    def setUp(self):
        self.DOCX_PATH = r'test.docx'
        self.paragraphs = wth.retrieve_paragraphs(self.DOCX_PATH)
        self.text = "Hello"
        self.groups = self.paragraphs.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")
        # convert_to_html(DOCX_PATH)

    def test_paragraph_text(self):
        self.assertEqual(wth.paragraph_text(self.text), "<p>Hello</p>")

    def test_bold_text(self):
        self.assertEqual(wth.bold_text(self.text), "<b>Hello</b>")

    def test_italicize_text(self):
        self.assertEqual(wth.italicize_text(self.text), "<i>Hello</i>")

    def test_underline_text(self):
        self.assertEqual(wth.underline_text(self.text), "<span style='text-decoration: underline;'>Hello</span>")

    def test_is_bold(self):
        pass

    def test_is_italics(self):
        pass

    def test_is_underline(self):
        pass

    def test_grab_text(self):
        pass

if __name__ == '__main__':
    unittest.main()