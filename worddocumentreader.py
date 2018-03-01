import zipfile
import xml.etree.ElementTree as ET

NAMESPACE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS_PARAGRAPHS = NAMESPACE + 'p'
NS_TEXT = NAMESPACE + 't'

def extract_xml_from_word(DOCX_PATH):
    docx = zipfile.ZipFile(DOCX_PATH)
    xml = ET.XML(docx.read('word/document.xml'))
    docx.close()
    return xml

def extract_paragraphs(xml):
    return xml.iter(NS_PARAGRAPHS)

def extract_text(paragraphs):
    result = []
    for paragraph in paragraphs:
        result.append("".join(p.text for p in paragraph.iter(NS_TEXT) if p.text != None))
    return result

def extract_text_from_path(path):
    return extract_text(extract_paragraphs(extract_xml_from_word(path)))