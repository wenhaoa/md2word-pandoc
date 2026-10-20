#!/usr/bin/env python
"""Check OOXML integrity and preservation of body media, tables and math."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import zipfile
from collections import Counter

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'


def snapshot(path):
    doc = Document(path)
    body = doc.element.body
    structures = ('f', 'rad', 'nary', 'm', 'eqArr', 'd', 'limLow', 'limUpp', 'acc', 'bar', 'groupChr')
    complex_math = {name: len(list(body.iter(M + name))) for name in structures}
    with zipfile.ZipFile(path) as archive:
        if archive.testzip():
            raise ValueError('DOCX归档损坏')
        media = Counter(hashlib.sha256(archive.read(name)).hexdigest()
                        for name in archive.namelist() if name.startswith('word/media/'))
    return {'drawings': len(list(body.iter(qn('w:drawing')))),
            'pictures': len(list(body.iter(qn('w:pict')))),
            'tables': len(list(body.iter(qn('w:tbl')))),
            'display_math': len(list(body.iter(M + 'oMathPara'))),
            'complex_math': complex_math, 'media': dict(media)}


def validate(path, baseline=None, with_cover=False):
    doc = Document(path)
    body = doc.element.body
    errors = []
    starts = list(body.iter(qn('w:bookmarkStart')))
    names = [element.get(qn('w:name')) for element in starts]
    ids = [element.get(qn('w:id')) for element in starts]
    ends = [element.get(qn('w:id')) for element in body.iter(qn('w:bookmarkEnd'))]
    if len(ids) != len(set(ids)) or len(names) != len(set(names)) or Counter(ids) != Counter(ends):
        errors.append('书签重复或起止不配对')
    end_by_id = {element.get(qn('w:id')): element for element in body.iter(qn('w:bookmarkEnd'))}
    for start in starts:
        name = start.get(qn('w:name'), '')
        if not re.match(r'^(?:Fig|Tbl|Eq|Sec)_', name):
            continue
        end = end_by_id.get(start.get(qn('w:id')))
        if end is not None and (start.getparent().tag != qn('w:p') or start.getparent() is not end.getparent()):
            errors.append('编号书签跨段，起止配对可能已被合并破坏：' + name)
    for element in body.iter(qn('w:instrText')):
        match = re.search(r'\bREF\s+(\S+)', element.text or '')
        if match and match[1] not in names:
            errors.append('交叉引用缺少书签：' + match[1])
    depth = 0
    for element in body.iter(qn('w:fldChar')):
        kind = element.get(qn('w:fldCharType'))
        if kind == 'begin':
            depth += 1
        elif kind == 'end':
            depth -= 1
            if depth < 0:
                errors.append('域结束符无对应起始符')
    if depth:
        errors.append('域起止不配对')
    text = ''.join(element.text or '' for element in body.iter(qn('w:t')))
    if '{{TITLE}}' in text or '{{DATE_CN}}' in text or 'Error! Reference source not found' in text or '错误!未找到引用源' in text:
        errors.append('残留封面占位符或引用错误')
    actual = snapshot(path)
    if baseline:
        for key in ('drawings', 'pictures', 'tables', 'display_math'):
            if actual[key] < baseline[key] or (not with_cover and actual[key] != baseline[key]):
                errors.append(f'{key}数量变化：{baseline[key]} -> {actual[key]}')
        for name, count in baseline['complex_math'].items():
            if actual['complex_math'][name] != count:
                errors.append(f'数学结构{name}数量变化：{count} -> {actual["complex_math"][name]}')
        if Counter(baseline['media']) - Counter(actual['media']):
            errors.append('原始图片数据丢失')
    if errors:
        raise ValueError('；'.join(errors))
    return actual


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx')
    parser.add_argument('--snapshot')
    parser.add_argument('--baseline')
    parser.add_argument('--with-cover', action='store_true')
    args = parser.parse_args()
    if args.snapshot:
        Path(args.snapshot).write_text(json.dumps(snapshot(args.docx)), encoding='utf-8')
    else:
        baseline = json.loads(Path(args.baseline).read_text(encoding='utf-8')) if args.baseline else None
        validate(args.docx, baseline, args.with_cover)
        print('  DOCX结构与内容保全检查通过')


if __name__ == '__main__':
    main()
