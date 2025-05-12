import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile
NAMESPACE = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}
_NAMESPACE = {v: k for k, v in NAMESPACE.items()}
PATTERN = re.compile(r'xmlns:ns(\d+)="(.*?)"')


def get_sectPr(docx_path: str, section: int, print=lambda *a, **b: None):
    """
    Extract all <w:sectPr> children elements as **format of section** from a .docx file.
    """
    with ZipFile(docx_path, 'r') as zip:
        with zip.open('word/document.xml') as Xml:
            tree = ET.parse(Xml)
            root = tree.getroot()

            sectPr_elements = root.findall('.//w:sectPr', {'w': NAMESPACE['w']})

            S = [''.join([
                ET.tostring(child, encoding='unicode')
                for child in list(sectPr)
            ]) for sectPr in sectPr_elements]

            print(f"Found {len(S)} <w:sectPr> elements in {docx_path}, use section={section}:")
            S = S[section]
            matches = set(PATTERN.findall(S))
            for match in matches:
                i, url = match
                S = S.replace(f'ns{i}', _NAMESPACE[url])
            for url, To in _NAMESPACE.items():
                S = S.replace(f' xmlns:{To}="{url}"', '')
            print('get_sectPr:\t', S)
            return S
