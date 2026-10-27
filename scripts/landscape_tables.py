"""Isolate explicitly marked tables on landscape pages after cover composition."""
import argparse
from copy import deepcopy
import json
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import autofit_tables
from table_sections import section_for_block


def section_in_paragraph(element):
    if element.tag != qn('w:p'):
        return None
    props = element.find(qn('w:pPr'))
    return None if props is None else props.find(qn('w:sectPr'))


def next_page(section, continue_numbering=False):
    kind = section.find(qn('w:type'))
    if kind is None:
        kind = OxmlElement('w:type')
        # w:type precedes page size in CT_SectPr.
        size = section.find(qn('w:pgSz'))
        if size is None:
            section.append(kind)
        else:
            size.addprevious(kind)
    kind.set(qn('w:val'), 'nextPage')
    if continue_numbering:
        numbering = section.find(qn('w:pgNumType'))
        if numbering is not None:
            numbering.attrib.pop(qn('w:start'), None)


def landscape(section):
    size = section.find(qn('w:pgSz'))
    if size is None:
        raise ValueError('横向表格所在节缺少纸张尺寸')
    width, height = int(size.get(qn('w:w'))), int(size.get(qn('w:h')))
    size.set(qn('w:w'), str(max(width, height)))
    size.set(qn('w:h'), str(min(width, height)))
    size.set(qn('w:orient'), 'landscape')


def boundary(section):
    paragraph = OxmlElement('w:p')
    props = OxmlElement('w:pPr')
    # A section boundary needs a paragraph, but should not leave a blank line.
    spacing = OxmlElement('w:spacing')
    for key, value in [('before', '0'), ('after', '0'), ('line', '20'), ('lineRule', 'exact')]:
        spacing.set(qn('w:' + key), value)
    props.append(spacing)
    props.append(section)
    paragraph.append(props)
    return paragraph


def has_content(element):
    if element.tag == qn('w:tbl'):
        return True
    if element.tag == qn('w:p'):
        return any(item.tag != qn('w:pPr') for item in element)
    return False


def isolate(doc, table):
    body = doc.element.body
    element = table._tbl
    if element.getparent() is not body:
        raise ValueError('横向标记仅支持正文表格，不支持单元格内表格')
    original = section_for_block(doc, element)
    before = []
    for sibling in element.itersiblings(preceding=True):
        if section_in_paragraph(sibling) is not None:
            break
        before.append(sibling)
    has_before = any(has_content(item) for item in before)
    if has_before:
        element.addprevious(boundary(deepcopy(original)))

    after = []
    for sibling in element.itersiblings():
        after.append(sibling)
        if sibling is original or section_in_paragraph(sibling) is original:
            break
    if any(has_content(item) for item in after):
        end = deepcopy(original)
        landscape(end)
        next_page(end, continue_numbering=has_before)
        element.addnext(boundary(end))
        # The following section keeps its original orientation and margins.
        next_page(original, continue_numbering=True)
    else:
        # A last/only table should not create an empty portrait trailing page.
        landscape(original)
        next_page(original, continue_numbering=has_before)

    for row in table.rows:
        props = row._tr.get_or_add_trPr()
        if props.find(qn('w:cantSplit')) is None:
            props.append(OxmlElement('w:cantSplit'))
    if table.rows:
        props = table.rows[0]._tr.get_or_add_trPr()
        if props.find(qn('w:tblHeader')) is None:
            props.append(OxmlElement('w:tblHeader'))


def apply(path, plan):
    if not plan['landscape']:
        return
    doc = Document(path)
    elements = list(doc.element.body.iter(qn('w:tbl')))
    offset = len(elements) - plan['table_count']
    if offset < 0:
        raise ValueError('正文表格数量少于布局计划')
    if len(set(plan['landscape'])) != len(plan['landscape']):
        raise ValueError('横向表格索引重复')
    tables = {table._tbl: (index, table) for index, table in enumerate(doc.tables)}
    selected = []
    # Validate every target before mutating anything.
    for index in plan['landscape']:
        if not isinstance(index, int) or not 0 <= index < plan['table_count']:
            raise ValueError('横向表格索引越界')
        element = elements[offset + index]
        if element not in tables:
            raise ValueError('横向目标不是正文顶层表格')
        selected.append(tables[element])
    for _, table in selected:
        isolate(doc, table)
    doc.save(path)
    # Refit only marked tables, using their actual section width.
    autofit_tables.run(path, body_only=True, table_indices={index for index, _ in selected}, compact=True)
    print('  已设置%d张表格单独横向，表后恢复原页面方向' % len(selected))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx')
    parser.add_argument('--plan', required=True)
    args = parser.parse_args()
    apply(args.docx, json.loads(Path(args.plan).read_text(encoding='utf-8')))
