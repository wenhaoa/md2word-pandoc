#!/usr/bin/env python
"""Number display equations and link references without flattening OMML."""
import argparse
import json
from pathlib import Path
import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx_safety import bookmark, field, next_bookmark_id, paragraphs, replace_run_refs, text_run

M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def m(tag):
    return '{%s}%s' % (M, tag)


def heading1_ids(doc):
    # Retain XML elements themselves; temporary Python proxy ids are unstable.
    return {p._element for p in doc.paragraphs if p.style and p.style.name.lower() == 'heading 1'}


def has_display_math(element):
    return bool(element.findall('.//' + m('oMathPara')))


def format_equation_para(paragraph, width):
    properties = paragraph._p.get_or_add_pPr()
    for name in ('jc', 'ind', 'tabs', 'spacing'):
        for old in properties.findall(qn('w:' + name)):
            properties.remove(old)
    align = OxmlElement('w:jc')
    align.set(qn('w:val'), 'left')
    properties.append(align)
    indent = OxmlElement('w:ind')
    for name in ('firstLine', 'left', 'right'):
        indent.set(qn('w:' + name), '0')
    properties.append(indent)
    spacing = OxmlElement('w:spacing')
    for name, value in (('line', '360'), ('lineRule', 'auto'), ('beforeLines', '50'), ('afterLines', '50')):
        spacing.set(qn('w:' + name), value)
    properties.append(spacing)
    tabs = OxmlElement('w:tabs')
    for kind, position in (('center', int(width / 2)), ('right', int(width))):
        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), kind)
        tab.set(qn('w:pos'), str(position))
        tabs.append(tab)
    properties.append(tabs)


def tab_run():
    run = OxmlElement('w:r')
    run.append(OxmlElement('w:tab'))
    return run


def run(path, dry=False, plan=None):
    doc = Document(path)
    headings = heading1_ids(doc)
    display = [p for p in paragraphs(doc) if has_display_math(p._p)]
    if plan is not None and len(plan) != len(display):
        raise ValueError(f'公式计划{len(plan)}项与显示公式{len(display)}个不一致')
    section = doc.sections[-1]
    width = (section.page_width - section.left_margin - section.right_margin) / 635
    chapter, numbered, identifier = 0, 0, next_bookmark_id(doc)
    counts, marks = {}, {}
    for paragraph in paragraphs(doc):
        element = paragraph._p
        if element in headings:
            chapter += 1
        if not has_display_math(element):
            continue
        if chapter < 1:
            raise ValueError('显示公式之前没有一级章节标题')
        counts[chapter] = counts.get(chapter, 0) + 1
        sequence = counts[chapter]
        expected = f'{chapter}-{sequence}'
        key = (plan[numbered] if plan is not None else None) or expected
        if key != expected:
            raise ValueError(f'公式源编号{key}与所在章及顺序{expected}不一致；保留源文档并修正编号后再转换')
        name = 'Eq_' + key.replace('-', '_')
        if key in marks:
            raise ValueError('重复公式编号：' + key)
        marks[key] = name
        numbered += 1
        # Idempotent: existing equation fields are never appended a second time.
        instructions = ''.join(t.text or '' for t in element.iter(qn('w:instrText')))
        if 'SEQ 公式' in instructions or dry:
            continue
        format_equation_para(paragraph, width)
        element.find(qn('w:pPr')).addnext(tab_run())
        number_runs = [tab_run(), text_run('(')]
        number_runs += field(' STYLEREF 1 \\s ', str(chapter))
        number_runs.append(text_run('-'))
        number_runs += field(' SEQ 公式 \\* ARABIC \\s 1 ', str(sequence))
        number_runs.append(text_run(')'))
        for item in number_runs:
            element.append(item)
        # Parentheses remain outside the bookmark so references never double them.
        bookmark(element, number_runs[2], number_runs[-2], name, identifier)
        identifier += 1
    converted = 0
    pattern = re.compile(r'(?:公式|式)\s*[（(]?\s*(\d+)\s*[-–]\s*(\d+)\s*[）)]?')
    if not dry:
        def resolve(match, source):
            key = f'{match[1]}-{match[2]}'
            if key not in marks:
                return None
            return [text_run('式(', source)] + field(' REF %s \\h ' % marks[key], key, source) + [text_run(')', source)]
        for paragraph in paragraphs(doc):
            if not has_display_math(paragraph._p):
                converted += replace_run_refs(paragraph, pattern, resolve)
        doc.save(path)
    print(f'  公式{numbered}个，正文引用域化{converted}处' + ('（dry-run）' if dry else ''))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx')
    parser.add_argument('--dry-run', action='store_true')
    group = parser.add_mutually_exclusive_group()
    group.add_argument('--plan', help='逗号分隔的源编号')
    group.add_argument('--plan-file', help='JSON数组，每个显示公式一项；无源编号用null')
    args = parser.parse_args()
    plan = json.loads(Path(args.plan_file).read_text(encoding='utf-8')) if args.plan_file else (args.plan.split(',') if args.plan else None)
    run(args.docx, args.dry_run, plan)


if __name__ == '__main__':
    main()
