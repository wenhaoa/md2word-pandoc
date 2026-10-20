#!/usr/bin/env python
"""Create unique heading bookmarks and preserve text while linking references."""
import argparse
import re
from docx import Document
from docx.oxml.ns import qn
from docx_safety import bookmark, field, next_bookmark_id, paragraphs, replace_run_refs, text_run


def level_of(paragraph):
    match = re.fullmatch(r'heading (\d+)', paragraph.style.name.lower()) if paragraph.style else None
    return int(match[1]) if match else None


def run(path, dry=False):
    doc = Document(path)
    counts, marks = [0] * 9, {}
    identifier = next_bookmark_id(doc)
    for paragraph in doc.paragraphs:
        level = level_of(paragraph)
        if level is None:
            continue
        counts[level - 1] += 1
        counts[level:] = [0] * (9 - level)
        number = '.'.join(str(n) for n in counts[:level])
        name = 'Sec_' + number.replace('.', '_')
        marks[number] = name
        if not dry:
            children = [child for child in paragraph._p if child.tag != qn('w:pPr')]
            if children:
                bookmark(paragraph._p, children[0], children[-1], name, identifier)
                identifier += 1
    pattern = re.compile(r'第\s*(\d+(?:\.\d+)*)\s*([章节])')
    def resolve(match, source):
        key, kind = match.groups()
        if key not in marks or (kind == '章' and '.' in key) or (kind == '节' and '.' not in key):
            return None
        return [text_run('第', source)] + field(' REF %s \\r \\h ' % marks[key], key, source) + [text_run(kind, source)]
    converted = 0
    if not dry:
        for paragraph in paragraphs(doc):
            if not level_of(paragraph):
                converted += replace_run_refs(paragraph, pattern, resolve)
        doc.save(path)
    print(f'  标题书签{len(marks)}个；章节引用域化{converted}处')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx')
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()
    run(args.docx, args.dry_run)


if __name__ == '__main__':
    main()
