"""Resolve the section governing a block (section properties end the section)."""
from docx.oxml.ns import qn


def section_for_block(doc, block):
    body = doc.element.body
    while block.getparent() is not body:
        block = block.getparent()
        if block is None:
            raise ValueError('表格不属于正文')
    for item in [block, *block.itersiblings()]:
        if item.tag == qn('w:sectPr'):
            return item
        if item.tag == qn('w:p'):
            props = item.find(qn('w:pPr'))
            if props is not None:
                section = props.find(qn('w:sectPr'))
                if section is not None:
                    return section
    raise ValueError('正文缺少节属性')
