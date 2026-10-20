#!/usr/bin/env python
"""Resolve global/chapter figure and table references from caption bookmarks."""
import argparse
import re
from docx import Document
from docx.oxml.ns import qn
from docx_safety import field, paragraphs, replace_run_refs, text_run


def md_caption_numbers(md_path):
    figures, tables, fence = [], [], None
    with open(md_path, encoding='utf-8') as stream:
        for line in stream:
            token = re.match(r'^\s*(`{3,}|~{3,})', line)
            if token:
                if fence is None:
                    fence = token[1][0]
                elif token[1][0] == fence:
                    fence = None
                continue
            if fence or line.startswith(('    ', '\t')):
                continue
            text = line.strip()
            if text.startswith('**') and text.endswith('**'):
                text = text[2:-2].strip()
            match = re.fullmatch(r'(图|表)\s*(\d+(?:[-.]\d+)?)\s*(\D.*)', text)
            if match:
                (figures if match[1] == '图' else tables).append(match[2].replace('.', '-'))
    return figures, tables


def run(docx_path, md_path=None, dry=False):
    doc = Document(docx_path)
    maps = {'图': {}, '表': {}}
    for element in doc.element.body.iter(qn('w:bookmarkStart')):
        name = element.get(qn('w:name'), '')
        match = re.fullmatch(r'(Fig|Tbl)_(\d+(?:_\d+)?)', name)
        if match:
            maps['图' if match[1] == 'Fig' else '表'][match[2].replace('_', '-')] = name
    if md_path:
        figures, tables = md_caption_numbers(md_path)
        for kind, values in (('图', figures), ('表', tables)):
            missing = set(values) - maps[kind].keys()
            if missing:
                raise ValueError('题注未生成编号书签：' + kind + '、'.join(sorted(missing)))
    pattern = re.compile(r'(图|表)(\d+(?:[-.]\d+)?)(?!\d|[-–.]\d)')
    def resolve(match, source):
        kind, key = match.groups()
        key = key.replace('.', '-')
        if key not in maps[kind]:
            return None
        return [text_run(kind, source)] + field(' REF %s \\h ' % maps[kind][key], key, source)
    converted = 0
    if not dry:
        for paragraph in paragraphs(doc):
            instructions = ''.join(t.text or '' for t in paragraph._p.iter(qn('w:instrText')))
            if 'SEQ 图' in instructions or 'SEQ 表' in instructions:
                continue
            converted += replace_run_refs(paragraph, pattern, resolve)
        doc.save(docx_path)
    print(f"  题注书签：图{len(maps['图'])}个、表{len(maps['表'])}个；正文引用域化{converted}处")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx')
    parser.add_argument('--md')
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()
    run(args.docx, args.md, args.dry_run)


if __name__ == '__main__':
    main()
