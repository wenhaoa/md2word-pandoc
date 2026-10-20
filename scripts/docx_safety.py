"""Shared OOXML helpers for lossless numbering and reference edits."""
from copy import deepcopy
import re
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def paragraphs(doc):
    return [Paragraph(el, doc) for el in doc.element.body.iter(qn('w:p'))]


def next_bookmark_id(doc):
    return max((int(el.get(qn('w:id'), '0')) for el in doc.element.body.iter(qn('w:bookmarkStart'))), default=0) + 1


def bookmark(parent, first, last, name, identifier):
    existing = parent.xpath('.//w:bookmarkStart[@w:name="%s"]' % name)
    if existing:
        return
    start, end = OxmlElement('w:bookmarkStart'), OxmlElement('w:bookmarkEnd')
    start.set(qn('w:name'), name)
    start.set(qn('w:id'), str(identifier))
    end.set(qn('w:id'), str(identifier))
    first.addprevious(start)
    last.addnext(end)


def text_run(text, source=None):
    run = OxmlElement('w:r')
    if source is not None:
        properties = source.find(qn('w:rPr'))
        if properties is not None:
            run.append(deepcopy(properties))
    element = OxmlElement('w:t')
    element.set(qn('xml:space'), 'preserve')
    element.text = text
    run.append(element)
    return run


def field(instr, cached='?', source=None):
    result = []
    for kind, text, code in (('begin', None, None), (None, None, instr),
                              ('separate', None, None), (None, cached, None), ('end', None, None)):
        run = text_run(text, source) if text is not None else OxmlElement('w:r')
        if kind:
            element = OxmlElement('w:fldChar')
            element.set(qn('w:fldCharType'), kind)
            # Cached source numbers are already valid. Marking every REF dirty
            # makes Word prompt for field updates before the document opens.
            run.append(element)
        if code:
            element = OxmlElement('w:instrText')
            element.set(qn('xml:space'), 'preserve')
            element.text = code
            run.append(element)
        result.append(run)
    return result


def replace_run_refs(paragraph, pattern, resolve):
    """Replace only ordinary text runs; keep run properties and field contents."""
    converted = 0
    if paragraph.style and paragraph.style.name.lower() in ('source code', 'verbatim', 'code'):
        return 0
    depth = 0
    for run in list(paragraph.runs):
        element = run._r
        markers = element.findall(qn('w:fldChar'))
        if markers:
            for marker in markers:
                kind = marker.get(qn('w:fldCharType'))
                depth += 1 if kind == 'begin' else (-1 if kind == 'end' else 0)
            continue
        if depth or any(ch.tag not in (qn('w:rPr'), qn('w:t')) for ch in element):
            continue
        if run.style and run.style.name.lower() in ('verbatim char', 'source code', 'code'):
            continue
        value = run.text or ''
        pieces, position = [], 0
        for match in pattern.finditer(value):
            replacement = resolve(match, element)
            if replacement is None:
                continue
            if match.start() > position:
                pieces.append(text_run(value[position:match.start()], element))
            pieces.extend(replacement)
            position = match.end()
            converted += 1
        if position:
            if position < len(value):
                pieces.append(text_run(value[position:], element))
            for piece in pieces:
                element.addprevious(piece)
            element.getparent().remove(element)
    return converted
