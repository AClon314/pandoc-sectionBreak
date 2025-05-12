#!/bin/env python
import re
import os
import locale
import panflute as pan
DEBUG = os.environ.get("DEBUG", None)
if DEBUG:
    def Log(*args, **kwargs):
        """Same as print, but prints to stderr (which is not intercepted by Pandoc)."""
        pan.debug(*args, **kwargs)
else:
    Log = lambda *args, **kwargs: None
TAG = 'br '
HTML_CSS = """
<style>
@media print {
break {
    break-before: page;
}}
</style>
"""
PATTERN_CLASS = re.compile(r'(?<=<br )\w+')
CLASS_TO_XML = {
    'section': '<w:p><w:pPr><w:sectPr>{}</w:sectPr></w:pPr></w:p>',
    'continue': '<w:p><w:pPr><w:sectPr><w:type w:val="continuous" />{}</w:sectPr></w:pPr></w:p>',
    'odd': '<w:p><w:pPr><w:sectPr><w:type w:val="oddPage" />{}</w:sectPr></w:pPr></w:p>',
    'even': '<w:p><w:pPr><w:sectPr><w:type w:val="evenPage" />{}</w:sectPr></w:pPr></w:p>',
    'col': '<w:r><w:rPr>{}</w:rPr><w:br w:type="column" /></w:r>',
}
CLASS_TO_ELEM: dict[str, pan.Element] = {}
IS_DOC_XML = None
_LANG = locale.getdefaultlocale()[0]
LANG = _LANG if _LANG else 'en_US'
I18N = {
    'warn_no_template': {
        'en_US': '⚠️ Not found `template: {temp}` as reference-doc in input metadata, you may have to adjust the format of new section, which fallback to pandoc default docx!',
        'zh_CN': '⚠️ 在元数据中未找到`template: {temp}`字段，您可能要调整新节的页边距等格式，回退到pandoc原生docx模板！',
    }
}


def get_metadata(doc: pan.Doc, key: str, default: str | None = None):
    return pan.stringify(doc.metadata[key]) if key in doc.metadata else default


def docx(elem: pan.RawInline):
    """docx filter"""
    classes = PATTERN_CLASS.findall(elem.text)
    for cls in classes:
        if cls in CLASS_TO_ELEM.keys():
            _elem = CLASS_TO_ELEM[cls]
            Log('docx:\t', getattr(_elem, 'text', 'Error'))
            return _elem


def action(elem: pan.Element, doc: pan.Doc | None = None):
    # Log(type(elem), str(elem)[:128])
    if IS_DOC_XML and isinstance(elem, pan.Para):
        if len(elem.content) == 1:
            inline = elem.content[0]
            if isinstance(inline, pan.RawInline) and TAG in inline.text:
                return docx(inline)
    return elem


def prepare(doc: pan.Doc):
    global IS_DOC_XML, CLASS_TO_XML, CLASS_TO_ELEM
    if 'doc' in doc.format:
        IS_DOC_XML = True
        template: str = get_metadata(doc, 'template', 'reference-doc.docx')  # type: ignore
        section = int(get_metadata(doc, 'section', '0'))    # type: ignore
        template = os.path.join(os.getcwd(), template)
        if template and os.path.exists(template):
            from pandoc_sectionBreak.docx_parse import get_sectPr
            sect_format = get_sectPr(template, section=section, print=Log)
            CLASS_TO_ELEM = {k: pan.RawBlock(v.format(sect_format), format='openxml') for k, v in CLASS_TO_XML.items()}
        else:
            pan.debug(I18N['warn_no_template'][LANG].format(temp=template))
    elif doc.format == 'html' or 'markdown' in doc.format:
        doc.content.insert(0, pan.RawBlock(HTML_CSS, format='html'))
    else:
        pan.debug(f'Skip un-implemented doc format: {getattr(doc, 'format', 'None')}')


def main(doc: pan.Doc | None = None):
    return pan.run_filters(prepare=prepare, actions=[action], doc=doc)


if __name__ == '__main__':
    main()
