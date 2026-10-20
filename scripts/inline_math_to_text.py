# -*- coding: utf-8 -*-
"""
inline_math_to_text.py - 把 docx 里的“简单行内公式”降级为 普通文本 + 上下标

规则:
  - 只处理行内 m:oMath（显示公式 m:oMathPara 一律不动）
  - 仅当公式由 m:r / m:sSub / m:sSup / m:sSubSup 组成时才转换
  - 出现分式、根式、积分、算子、矩阵等复杂结构时，保留原公式

用法:
    python inline_math_to_text.py <docx> [--dry-run]
"""

import sys
import re
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
def m(tag):
    return '{%s}%s' % (M, tag)

SIMPLE = {m('r'), m('sSub'), m('sSup'), m('sSubSup')}


def om_text(el):
    """取 OMML 元素内的纯文本；乘号点统一为 U+00B7（N·m 的习惯写法）。"""
    txt = ''.join(t.text or '' for t in el.findall('.//' + m('t')))
    return txt.replace('\u22c5', '\u00b7')


def decompose(om):
    """把行内 OMML 拆成 [(base, sub, sup), ...]；遇到复杂结构返回 None。"""
    allowed = SIMPLE | {m('oMath'), m('e'), m('sub'), m('sup'), m('t'),
                        m('sSubPr'), m('sSupPr'), m('sSubSupPr')}
    # Inspect the entire tree, including scripts. Flattening only the top level
    # destroys nested fractions and styled mathematical symbols.
    if any(element.tag not in allowed for element in om.iter()):
        return None
    for element in om.iter():
        if element.tag in (m('e'), m('sub'), m('sup')):
            if any(child.tag != m('r') for child in element):
                return None
    parts = []
    for ch in list(om):
        if ch.tag not in SIMPLE:
            return None
        if ch.tag == m('r'):
            parts.append((om_text(ch), None, None))
            continue
        e = ch.find(m('e'))
        sub = ch.find(m('sub'))
        sup = ch.find(m('sup'))
        base = om_text(e) if e is not None else ''
        s = om_text(sub) if sub is not None else None
        p = om_text(sup) if sup is not None else None
        parts.append((base, s, p))
    # 丢掉全空片段
    parts = [x for x in parts if (x[0] or x[1] or x[2])]
    return parts or None


def make_run(text, vert=None, italic=False):
    """生成一个 w:r。vert='subscript'/'superscript'。"""
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    if italic:
        rPr.append(OxmlElement('w:i'))
    if vert:
        va = OxmlElement('w:vertAlign')
        va.set(qn('w:val'), vert)
        rPr.append(va)
    if len(rPr):
        r.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    return r


def looks_like_var(s):
    """单个拉丁字母（含希腊字母转写）视为变量，按惯例用斜体。"""
    return bool(re.fullmatch(r'[A-Za-z\u0391-\u03c9]', s or ''))


def convert(doc, dry_run=False):
    body = doc.element.body
    inline = [e for e in body.iter(m('oMath'))]
    # 排除显示公式内部的行内 oMath（m:oMathPara 的子元素）
    def in_display(el):
        p = el.getparent()
        while p is not None:
            if p.tag == m('oMathPara'):
                return True
            p = p.getparent()
        return False

    conv = 0
    keep = 0
    for om in inline:
        if in_display(om):
            continue
        parts = decompose(om)
        if parts is None:
            keep += 1
            continue
        if dry_run:
            conv += 1
            continue
        parent = om.getparent()
        idx = list(parent).index(om)
        new_runs = []
        for base, sub, sup in parts:
            if base:
                new_runs.append(make_run(base, italic=looks_like_var(base)))
            if sub:
                new_runs.append(make_run(sub, vert='subscript'))
            if sup:
                new_runs.append(make_run(sup, vert='superscript'))
        for r in new_runs:
            parent.insert(idx, r)
            idx += 1
        parent.remove(om)
        conv += 1
    return conv, keep


def main():
    argv = [a for a in sys.argv[1:] if not a.startswith('--')]
    dry = '--dry-run' in sys.argv
    if not argv:
        print('Usage: python inline_math_to_text.py <docx> [--dry-run]')
        sys.exit(1)
    path = argv[0]
    doc = Document(path)
    conv, keep = convert(doc, dry_run=dry)
    if not dry:
        doc.save(path)
    print(f'  {conv} 个行内公式转为文本+上下标，{keep} 个复杂公式保留为公式')
    if dry:
        print('  (dry-run，未保存)')


if __name__ == '__main__':
    main()
