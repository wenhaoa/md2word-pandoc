"""
merge_cover.py - 将封面+目录前缀模板与 Pandoc 生成的正文合并

用法:
    python merge_cover.py <模板docx> <正文docx> <输出docx> [--title "标题"]

依赖:
    pip install docxcompose
"""

import re
import datetime
import sys
import argparse
from docxcompose.composer import Composer
from docx import Document
from docx.oxml.ns import qn


def renumber_bookmark_pairs(doc, start_id=0):
    """按原ID同步重排书签两端，保留Pandoc章节与编号书签的嵌套关系。"""
    mapping = {}
    starts = list(doc.element.body.iter(qn('w:bookmarkStart')))
    ends = list(doc.element.body.iter(qn('w:bookmarkEnd')))
    for element in starts:
        old_id = element.get(qn('w:id'))
        if old_id in mapping:
            raise ValueError('书签起始ID重复：' + str(old_id))
        mapping[old_id] = str(start_id + len(mapping))
    for element in ends:
        old_id = element.get(qn('w:id'))
        if old_id not in mapping:
            raise ValueError('书签结束ID没有对应起始：' + str(old_id))
    for element in starts + ends:
        element.set(qn('w:id'), mapping[element.get(qn('w:id'))])
    return start_id + len(mapping)


class BookmarkPreservingComposer(Composer):
    def renumber_bookmarks(self):
        # docxcompose分别按出现顺序给起止端编号，会破坏嵌套书签配对。
        renumber_bookmark_pairs(self.doc)


def replace_cover_placeholders(doc, replacements):
    """替换封面占位符，兼容被Word拆分到多个run的标题和年月。"""
    for para in doc.paragraphs:
        matched = {key: value for key, value in replacements.items() if key in para.text}
        if not matched:
            continue
        for placeholder, value in matched.items():
            replaced = False
            for run in para.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)
                    replaced = True
            if replaced:
                continue
            new_text = para.text.replace(placeholder, value)
            if para.runs:
                para.runs[0].text = new_text
                for run in para.runs[1:]:
                    run.text = ''


def replace_title_placeholder(doc, title):
    """替换模板中的 {{TITLE}} 占位符为实际标题。

    WHY: Word 可能将 {{TITLE}} 拆分到多个 run 中，
    因此先尝试 run 级替换，若失败则回退到段落级重组。
    """
    replace_cover_placeholders(doc, {'{{TITLE}}': title})


def add_table_borders(doc):
    """给文档中所有表格添加全框线（单实线）并居中对齐。

    WHY: Pandoc 生成的 docx 表格默认无边框且左对齐，不使用模板的 Table Normal 样式，
    只能通过后处理在 OOXML 层面添加 tblBorders 和 jc 元素。
    """
    from docx.oxml.ns import qn
    from lxml import etree

    # 6 条边框：上、下、左、右、内部横线、内部竖线
    BORDER_NAMES = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']
    # 边框属性：单实线，0.5pt（4 half-points），黑色
    BORDER_ATTRS = {'val': 'single', 'sz': '4', 'space': '0', 'color': '000000'}

    for table in doc.tables:
        tbl = table._tbl
        tbl_pr = tbl.tblPr
        if tbl_pr is None:
            tbl_pr = etree.SubElement(tbl, qn('w:tblPr'))

        # 移除已有的 tblBorders（避免重复）
        existing = tbl_pr.find(qn('w:tblBorders'))
        if existing is not None:
            tbl_pr.remove(existing)

        # 创建新的 tblBorders
        borders = etree.SubElement(tbl_pr, qn('w:tblBorders'))
        for name in BORDER_NAMES:
            border = etree.SubElement(borders, qn(f'w:{name}'))
            for attr, value in BORDER_ATTRS.items():
                border.set(qn(f'w:{attr}'), value)

        # WHY: Pandoc 默认生成左对齐表格，通过 jc 元素设置居中
        existing_jc = tbl_pr.find(qn('w:jc'))
        if existing_jc is not None:
            tbl_pr.remove(existing_jc)
        jc = etree.SubElement(tbl_pr, qn('w:jc'))
        jc.set(qn('w:val'), 'center')

def _remove_first_empty_paragraph(doc):
    """删除正文（第一个段落级分节符之后）的第一个空段落。

    WHY: docxcompose.append() 合并封面和正文时，会在分节符后
    残留一个空的 Normal 段落，导致正文第一行总是空行。
    """
    from docx.oxml.ns import qn

    body = doc.element.body
    children = list(body)

    # WHY: 找第一个段落级 sectPr（封面/正文分界符），
    #       不能用最后一个（body-level sectPr 是文档末尾属性，不是分节符）
    first_sect_idx = -1
    for i, c in enumerate(children):
        if c.tag == qn('w:p'):
            ppr = c.find(qn('w:pPr'))
            if ppr is not None and ppr.find(qn('w:sectPr')) is not None:
                first_sect_idx = i
                break

    if first_sect_idx < 0:
        return

    # Only an immediately adjacent, structurally empty paragraph is removable.
    # Drawings, math, fields, breaks and bookmarks are content even without w:t.
    index = first_sect_idx + 1
    if index >= len(children):
        return
    paragraph = children[index]
    if paragraph.tag != qn('w:p'):
        return
    for child in paragraph:
        if child.tag == qn('w:pPr'):
            if child.find(qn('w:sectPr')) is not None:
                return
            continue
        if child.tag != qn('w:r'):
            return
        for item in child:
            if item.tag == qn('w:rPr'):
                continue
            if item.tag != qn('w:t') or (item.text or '').strip():
                return
    body.remove(paragraph)
    print('   已删除正文开头空段落')


DATE_PATTERNS = [
    re.compile(r"^\s*\d{4}\s*年\s*\d{1,2}\s*月\s*$"),
    re.compile(r"^\s*\d{4}\s*[-/]\s*\d{1,2}\s*$"),
    re.compile(r"^\s*\d{4}\.\d{1,2}\s*$"),
]


def replace_cover_date(doc, date_text=None):
    """把封面上的日期改成本年月。

    WHY: 模板里的日期是字面文本（如"2026年3月"），不是占位符，
    每次转换都会沿用旧日期。这里在封面区（首个 Heading 1 之前）
    按日期形态识别并替换，不影响正文中的任何年份表述。
    """
    if date_text is None:
        today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
        date_text = "%d年%d月" % (today.year, today.month)

    replace_cover_placeholders(doc, {'{{DATE_CN}}': date_text})

    # 封面区 = 首个 Heading 1 之前的段落（按段落序，不用 body 索引）
    cover = []
    for p in doc.paragraphs:
        try:
            if p.style is not None and p.style.name.lower() == "heading 1":
                break
        except Exception:
            pass
        cover.append(p)

    n = 0
    for p in cover:
        txt = p.text.strip()
        if not any(pat.match(txt) for pat in DATE_PATTERNS):
            continue
        for r in p.runs:
            r.text = ""
        if p.runs:
            p.runs[0].text = date_text
        else:
            p.add_run(date_text)
        n += 1
    return n


def merge(prefix_path, body_path, output_path, title=None, date_text=None, date_cn=None):
    """合并前缀模板与正文文档。

    Args:
        prefix_path: 包含封面+目录的模板文件路径
        body_path: Pandoc 生成的正文文件路径
        output_path: 合并后的输出文件路径
        title: 可选，替换 {{TITLE}} 占位符的标题文本
    """
    date_text = date_text or date_cn
    # 1. 打开前缀模板
    prefix = Document(prefix_path)

    # 2. 替换标题占位符
    if title:
        replace_title_placeholder(prefix, title)

    # 封面日期改成本年月（模板里是字面日期，非占位符）
    if date_text:
        changed = replace_cover_date(prefix, date_text)
        print('   封面日期已更新为: %s' % date_text)

    # 模板自带字面日期时，即使未显式传入也刷新
    elif not date_text:
        changed = replace_cover_date(prefix, None)
        if changed:
            print('   封面日期已刷新为本年月')

    # 3. 给正文表格添加框线（在合并前处理，封面表格不受影响）
    body = Document(body_path)
    add_table_borders(body)

    # 4. 合并正文
    # 先隔离封面和正文的ID，再在合并时按原配对关系同步编号。
    next_id = renumber_bookmark_pairs(prefix)
    renumber_bookmark_pairs(body, next_id)
    composer = BookmarkPreservingComposer(prefix)
    composer.append(body)

    # 4.5 删除正文开头的空段落
    # WHY: docxcompose 合并时在分节符后会残留一个空 Normal 段落，
    #       导致正文第一行总是空行，需自动清除
    _remove_first_empty_paragraph(composer.doc)

    # 5. 保存
    composer.save(output_path)
    print(f"   合并完成: {output_path}")


def main():
    parser = argparse.ArgumentParser(description='合并封面模板与正文文档')
    parser.add_argument('prefix', help='封面+目录模板文件路径')
    parser.add_argument('body', help='正文文档路径')
    parser.add_argument('output', help='输出文件路径')
    parser.add_argument('--title', help='替换 {{TITLE}} 占位符的标题', default=None)
    parser.add_argument('--date-cn', help='指定封面中文年月，默认北京时间本年月', default=None)

    args = parser.parse_args()
    merge(args.prefix, args.body, args.output, args.title, date_cn=args.date_cn)


if __name__ == '__main__':
    main()
