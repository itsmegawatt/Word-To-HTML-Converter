import zipfile
import xml.etree.ElementTree as ET

NAMESPACE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS_PARAGRAPHS = NAMESPACE + 'p'
NS_TEXT = NAMESPACE + 't'
NS_GROUP = NAMESPACE + 'r'
NS_LINEBREAKS = NAMESPACE + "br"

DOCX_PATH = r'test.docx'

def extract_xml_from_word(DOCX_PATH):
    docx = zipfile.ZipFile(DOCX_PATH)
    xml = ET.XML(docx.read('word/document.xml'))
    docx.close()
    return xml

def get_text(group):
    return "".join(t.text for t in group.iter(NS_TEXT))

def replace_symbols(text):
    text = text.replace("…", "...")
    text = text.replace("—", "--")
    return text

def is_blank(text):
    return text == ''

def needs_linebreak_before(group):
    linebreaks = group.iter(NS_LINEBREAKS)
    needs_linebreak = False
    for linebreak in linebreaks:
        needs_linebreak = True
    return needs_linebreak

def extract_paragraphs(xml):
    return xml.iter(NS_PARAGRAPHS)

def extract_text(paragraphs):
    result = []
    for paragraph in paragraphs:
        final_text = ""
        groups = paragraph.iter(NS_GROUP)
        for group in groups:
            text = get_text(group)
            text = replace_symbols(text)
            if needs_linebreak_before(group):
                final_text += "\n"
            final_text += text
        if not is_blank(final_text):
            result.append(final_text)
    return result

def extract_text_from_path(path):
    return extract_text(extract_paragraphs(extract_xml_from_word(path)))