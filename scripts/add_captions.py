"""
add_captions.py - Pandoc docx 后处理：题注自动编号及表格、图片对齐

用法:
    python add_captions.py <docx文件> [--no-dedup]

功能:
    1. 识别 "图/表" 题注：已带 "图N-M" 的按原样式处理；不带编号的按紧邻的图片/表格判定为题注，
       两者都统一替换为 STYLEREF + SEQ 域，由 Word 自动编号。
    2. 保留全部同名题注；源编号重复时报错。
    3. 表格单元格：垂直居中 + 水平居中 + 清除缩进。
    4. 图片段落居中。

域代码结构:
    图{STYLEREF 1 \\s}-{SEQ 图 \\* ARABIC \\s 1} 标题

依赖:
    pip install python-docx lxml
"""

import sys
import re
import copy
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx_safety import bookmark, field, next_bookmark_id

# 题注：可带编号（图4-1）也可不带（图 系统架构）；分隔符兼容全/半角与破折号
CAPTION_PATTERN = re.compile(
    r'^\s*(图|表)\s*(?:([0-9]+(?:[\-.][0-9]+)?)\s*)?[、．.:：\-–—]?\s*(.*)$',
    re.S)
# 严格编号形式：只有这种形式无需紧邻证据即可判定为题注
NUMBERED_PATTERN = re.compile(r'^\s*(图|表)\s*[0-9]+(?:[-.][0-9]+)?\s*')
MAX_CAPTION_LENGTH = 80


def is_plausible_caption_title(title):
    """无编号题注的合理性判定。

    WHY: 正文里以“图/表”开头的句子（如“表 中数据为分析值，…”“图 中展示…”）
    常紧邻图片或表格，仅靠位置证据会误判为题注。题注标题通常短、无句号、无逗号，
    也不会以“中”“为”“是”等后接成分开头。
    """
    t = (title or '').strip()
    if not t or len(t) > 40:
        return False
    if t[-1] in '。.；;，,':
        return False
    if '，' in t or ',' in t:
        return False
    if t[0] in '中为是表图':
        return False
    return True


def make_field_runs(instr_text):
    """创建域代码 run 序列（begin + instrText + separate + result + end）。"""
    runs = []

    r_begin = OxmlElement('w:r')
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    r_begin.append(fld_begin)
    runs.append(r_begin)

    r_instr = OxmlElement('w:r')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = instr_text
    r_instr.append(instr)
    runs.append(r_instr)

    r_sep = OxmlElement('w:r')
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    r_sep.append(fld_sep)
    runs.append(r_sep)

    r_result = OxmlElement('w:r')
    t_result = OxmlElement('w:t')
    t_result.text = '?'
    r_result.append(t_result)
    runs.append(r_result)

    r_end = OxmlElement('w:r')
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    r_end.append(fld_end)
    runs.append(r_end)

    return runs


def apply_rpr_to_runs(run_elems, rpr_source_elem):
    """给 run 元素列表应用格式属性。"""
    if rpr_source_elem is None:
        return
    source_rpr = rpr_source_elem.find(qn('w:rPr'))
    if source_rpr is None:
        return
    for run_elem in run_elems:
        if run_elem.tag == qn('w:r'):
            new_rpr = copy.deepcopy(source_rpr)
            run_elem.insert(0, new_rpr)


def make_text_run(text):
    """创建文本 run。"""
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    return r


def _ensure_bold(run_elems):
    """给 run 元素列表强制添加加粗格式。"""
    for r in run_elems:
        if r.tag != qn('w:r'):
            continue
        rpr = r.find(qn('w:rPr'))
        if rpr is None:
            rpr = OxmlElement('w:rPr')
            r.insert(0, rpr)
        if rpr.find(qn('w:b')) is None:
            rpr.append(OxmlElement('w:b'))


def build_caption_runs(caption_type, rpr_source=None, number=None):
    """构建全域化题注 run 序列: 图{STYLEREF}-{SEQ} 标题。"""
    all_runs = []

    all_runs.append(make_text_run(caption_type))

    if number and '-' not in number:
        all_runs.extend(field(f' SEQ {caption_type} \\* ARABIC ', number))
    else:
        chapter, sequence = number.split('-') if number else ('?', '?')
        all_runs.extend(field(' STYLEREF 1 \\s ', chapter))
        all_runs.append(make_text_run('-'))
        all_runs.extend(field(f' SEQ {caption_type} \\* ARABIC \\s 1 ', sequence))

    all_runs.append(make_text_run(' '))

    if rpr_source is not None:
        apply_rpr_to_runs(all_runs, rpr_source)

    return all_runs


def set_caption_format(para, is_table=False):
    """题注段落：居中、无缩进；表题段前 0.5 行、图题段后 0.5 行。

    WHY: 用 w:beforeLines / w:afterLines（1/100 行为单位），
    “0.5 行”会随字号与行距自动换算。
    """
    p = para._element
    # 表题与下一张表保持同页，避免题注孤立在上一页页末。
    para.paragraph_format.keep_with_next = is_table
    para.paragraph_format.keep_together = True
    pPr = p.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p.insert(0, pPr)
    for tag in ('w:jc', 'w:ind', 'w:spacing'):
        for old in pPr.findall(qn(tag)):
            pPr.remove(old)
    ind = OxmlElement('w:ind')
    ind.set(qn('w:firstLine'), '0')
    ind.set(qn('w:left'), '0')
    pPr.append(ind)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    pPr.append(jc)
    sp = OxmlElement('w:spacing')
    if is_table:
        sp.set(qn('w:beforeLines'), '50')
    else:
        sp.set(qn('w:afterLines'), '50')
    pPr.append(sp)

def has_drawing(para_elem):
    """检查段落是否包含图片。"""
    return bool(para_elem.findall('.//' + qn('w:drawing')))


# ---------------------------------------------------------------- 表格居中
def center_table_cells(doc):
    """所有表格单元格：垂直居中 + 水平居中 + 清除缩进。

    WHY: Pandoc 生成的表格默认靠上、靠左对齐，多行文字时明显不整齐。
    垂直居中由 w:tcPr/w:vAlign 控制，水平居中由单元格内每个段落的 w:jc 控制。
    """
    cells = 0
    paras = 0
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                for old in tcPr.findall(qn('w:vAlign')):
                    tcPr.remove(old)
                vAlign = OxmlElement('w:vAlign')
                vAlign.set(qn('w:val'), 'center')
                tcPr.append(vAlign)
                cells += 1

                for p in cell.paragraphs:
                    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.first_line_indent = Pt(0)
                    p.paragraph_format.left_indent = Pt(0)
                    p.paragraph_format.right_indent = Pt(0)
                    paras += 1
    if cells:
        print(f'  {cells} 个单元格垂直居中、{paras} 个段落水平居中')


def center_images(doc):
    """将包含图片的段落设为居中对齐、无首行缩进。"""
    count = 0
    for p in doc.paragraphs:
        if has_drawing(p._element):
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.keep_with_next = True
            count += 1
    if count:
        print(f'  {count} 个图片段落居中')


# ---------------------------------------------------------------- 题注处理
def body_sequence(doc):
    """按文档顺序返回 body 直接子元素列表。"""
    return list(doc.element.body)


def prev_is_drawing_or_table(seq, idx):
    """前一个元素是图片段落或表格。"""
    for j in range(idx - 1, -1, -1):
        el = seq[j]
        if el.tag == qn('w:tbl'):
            return 'table'
        if el.tag == qn('w:p'):
            if has_drawing(el):
                return 'drawing'
            if (el.text_content().strip()
                    if hasattr(el, 'text_content') else True):
                return None
    return None


def next_nonempty_kind(seq, idx, window=3):
    """仅越过空段，遇到普通内容即停止。"""
    for j in range(idx + 1, min(len(seq), idx + 1 + window)):
        el = seq[j]
        if el.tag == qn('w:tbl'):
            return 'table'
        if el.tag == qn('w:p') and has_drawing(el):
            return 'drawing'
        if el.tag != qn('w:p') or any(child.tag != qn('w:pPr') for child in el):
            return None
    return None


def prev_nonempty_kind(seq, idx, window=3):
    """仅越过空段，遇到普通内容即停止。"""
    for j in range(idx - 1, max(-1, idx - 1 - window), -1):
        el = seq[j]
        if el.tag == qn('w:tbl'):
            return 'table'
        if el.tag == qn('w:p') and has_drawing(el):
            return 'drawing'
        if el.tag != qn('w:p') or any(child.tag != qn('w:pPr') for child in el):
            return None
    return None


def process_captions(doc, dedup=False):
    """识别并域化题注；同名题注保留，不凭标题文字删除内容。"""
    seq = body_sequence(doc)
    index_of = {id(el): i for i, el in enumerate(seq)}

    done = 0
    bm_id = next_bookmark_id(doc)
    chapter = 0
    counts = {}

    # 逐个段落处理（doc.paragraphs 与 body 中的 w:p 一一对应）
    for p in list(doc.paragraphs):
        if p.style and p.style.name.lower() == 'heading 1':
            chapter += 1
        if p.style and p.style.name.lower() in ('source code', 'verbatim', 'code'):
            continue
        instructions = ''.join(t.text or '' for t in p._element.iter(qn('w:instrText')))
        if 'SEQ 图' in instructions or 'SEQ 表' in instructions:
            continue
        text = p.text.strip()
        if not text or len(text) > MAX_CAPTION_LENGTH:
            continue

        m = CAPTION_PATTERN.match(text)
        if not m:
            continue

        caption_type = m.group(1)
        title_text = (m.group(3) or '').strip()
        if not title_text:
            continue

        numbered = bool(NUMBERED_PATTERN.match(text))
        idx = index_of.get(id(p._element))
        if idx is None:
            continue

        if not numbered:
            # 无编号：必须有紧邻证据，避免把正文里以"图/表"开头的句子当成题注
            if not is_plausible_caption_title(title_text):
                continue
            if caption_type == '图':
                if not (prev_nonempty_kind(seq, idx) == 'drawing'
                        or next_nonempty_kind(seq, idx) == 'drawing'):
                    continue
            else:
                if not (next_nonempty_kind(seq, idx) == 'table'
                        or prev_nonempty_kind(seq, idx) == 'table'):
                    continue

        # Rebuilding a text-only caption must not discard a hyperlink, equation,
        # image, bookmark or field embedded in the original paragraph.
        if any(child.tag not in (qn('w:pPr'), qn('w:r')) for child in p._element):
            continue
        if any(child.tag not in (qn('w:rPr'), qn('w:t'))
               for run in p._element.findall(qn('w:r')) for child in run):
            continue
        counts[(caption_type, chapter)] = counts.get((caption_type, chapter), 0) + 1
        number = (m.group(2) or f'{chapter}-{counts[(caption_type, chapter)]}').replace('.', '-')

        first_run_elem = p.runs[0]._element if p.runs else None

        for r in p._element.findall(qn('w:r')):
            p._element.remove(r)

        caption_runs = build_caption_runs(caption_type, first_run_elem, number)

        r_title = make_text_run(title_text)
        if first_run_elem is not None:
            source_rpr = first_run_elem.find(qn('w:rPr'))
            if source_rpr is not None:
                r_title.insert(0, copy.deepcopy(source_rpr))

        all_runs = caption_runs + [r_title]
        _ensure_bold(all_runs)

        ppr = p._element.find(qn('w:pPr'))
        insert_after = ppr if ppr is not None else None
        for run_elem in all_runs:
            if insert_after is not None:
                insert_after.addnext(run_elem)
                insert_after = run_elem
            else:
                p._element.append(run_elem)

        name = ('Fig_' if caption_type == '图' else 'Tbl_') + number.replace('-', '_')
        if any(el.get(qn('w:name')) == name for el in doc.element.body.iter(qn('w:bookmarkStart'))):
            raise ValueError(f'重复题注编号：{caption_type}{number}')
        bookmark(p._element, caption_runs[1], caption_runs[-2], name, bm_id)
        bm_id += 1

        set_caption_format(p, is_table=(caption_type == '表'))

        done += 1
        flag = '' if numbered else '（原无编号）'
        print(f'  ok {caption_type}{flag}: "{title_text[:30]}"')

    return done


def main():
    argv = [a for a in sys.argv[1:]]
    argv = [a for a in argv if not a.startswith('--')]
    if not argv:
        print('Usage: python add_captions.py <docx> [--no-dedup]')
        sys.exit(1)

    docx_path = argv[0]
    print(f'  Processing: {docx_path}')

    doc = Document(docx_path)

    count = process_captions(doc)
    center_table_cells(doc)
    center_images(doc)

    doc.save(docx_path)
    print(f'  {count} 条题注已转为自动编号域')
    print('  Tip: 在 Word 中 Ctrl+A 后按 F9 更新域即可显示编号')


if __name__ == '__main__':
    main()
