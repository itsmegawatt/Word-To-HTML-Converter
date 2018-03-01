import worddocumentreader as wdr

NAMESPACE_GROUP = wdr.NAMESPACE + "r"
NAMESPACE_BOLD = wdr.NAMESPACE + "b"
NAMESPACE_ITALICS = wdr.NAMESPACE + "i"
NAMESPACE_UNDERLINE = wdr.NAMESPACE + "u"


def retrieve_paragraphs(docx_path):
    xml = wdr.extract_xml_from_word(docx_path)
    return wdr.extract_paragraphs(xml)

def paragraph_text(text):
    return "<p>" + text + "</p>"

def bold_text(text):
    return "<b>" + text + "</b>"

def italicize_text(text):
    return "<i>" + text + "</i>"

def underline_text(text):
    return "<span style='text-decoration: underline;'>" + text + "</span>"


DOCX_PATH = r'test.docx'
paragraphs = retrieve_paragraphs(DOCX_PATH)
for paragraph in paragraphs:
    result = ""
    groups = paragraph.iter(NAMESPACE_GROUP)
    for group in groups:
        bolds = group.iter(NAMESPACE_BOLD)
        is_bold = False
        for bold in bolds:
            is_bold = True

        italics = group.iter(NAMESPACE_ITALICS)
        is_italic = False
        for italic in italics:
            is_italic = True

        underlines = group.iter(NAMESPACE_UNDERLINE)
        is_underline = False
        for underline in underlines:
            is_underline = True

        texts = group.iter(wdr.NS_TEXT)
        for text in texts:
            mini_result = text.text

        text = "".join(t.text for t in group.iter(wdr.NS_TEXT))
        if is_bold:
            text = bold_text(text)
        if is_italic:
            text = italicize_text(text)
        if is_underline:
            text = underline_text(text)
        result += text
    print(result)