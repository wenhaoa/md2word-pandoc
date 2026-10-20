# -*- coding: utf-8 -*-
"""format_citations.py - 把正文中的 [1,2] / [3] 类参考文献引用改为上标。"""
import re, sys
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx_safety import paragraphs, replace_run_refs, text_run

CITE = re.compile(r"\[(\d+(?:[,\-\u2013]\d+)*)\]")


def mksup(text):
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    va = OxmlElement("w:vertAlign"); va.set(qn("w:val"), "superscript")
    rPr.append(va); r.append(rPr)
    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text
    r.append(t)
    return r


def run(path):
    doc = Document(path)
    n = 0
    references = False
    for p in paragraphs(doc):
        if p.style and p.style.name.lower().startswith('heading '):
            references = bool(re.search(r'参考文献|references', p.text, re.I))
        if references or re.match(r'^\s*\[\d+\]', p.text):
            continue
        def resolve(match, source):
            props = source.find(qn('w:rPr'))
            if props is not None and props.find(qn('w:vertAlign')) is not None:
                return None
            result = text_run(match[0], source)
            props = result.find(qn('w:rPr'))
            if props is None:
                props = OxmlElement('w:rPr')
                result.insert(0, props)
            align = OxmlElement('w:vertAlign')
            align.set(qn('w:val'), 'superscript')
            props.append(align)
            return [result]
        n += replace_run_refs(p, CITE, resolve)
    doc.save(path)
    print("  参考文献引用转上标 %d 处" % n)


if __name__ == "__main__":
    run(sys.argv[1])
