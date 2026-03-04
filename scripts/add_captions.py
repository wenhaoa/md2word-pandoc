"""
add_captions.py - 将图表题注替换为 Word SEQ 域自动编号 + 调整图片尺寸

用法:
    python add_captions.py <docx文件>

功能:
    1. 扫描 Pandoc 生成的 docx，识别 "图N-M" / "表N-M" 题注并替换为域
    2. 将嵌入图片宽度调整为页面宽度

域代码结构:
    图{STYLEREF "Heading 1" \\s}-{SEQ 图 \\* ARABIC \\s 1} 标题

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
from docx.shared import Pt, Cm, Emu


# 匹配独立题注行
CAPTION_PATTERN = re.compile(r'^(图|表)\s*\d+[-\.]\d+\s*')
MAX_CAPTION_LENGTH = 60

# 图片目标宽度（A4 页面可用宽度约 15.9cm，取 14.5cm 留边距）
IMAGE_TARGET_WIDTH_CM = 14.5


def make_field_runs(instr_text):
    """创建域代码 run 序列（begin + instrText + separate + result + end）。

    Args:
        instr_text: 域指令文本，如 ' SEQ 图 \\* ARABIC \\s 1 '
    """
    runs = []

    # begin
    r_begin = OxmlElement('w:r')
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    r_begin.append(fld_begin)
    runs.append(r_begin)

    # instrText
    r_instr = OxmlElement('w:r')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = instr_text
    r_instr.append(instr)
    runs.append(r_instr)

    # separate
    r_sep = OxmlElement('w:r')
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    r_sep.append(fld_sep)
    runs.append(r_sep)

    # result placeholder
    r_result = OxmlElement('w:r')
    t_result = OxmlElement('w:t')
    t_result.text = '?'
    r_result.append(t_result)
    runs.append(r_result)

    # end
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


def build_caption_runs(caption_type, rpr_source=None):
    """构建全域化题注 run 序列。

    生成: 图{STYLEREF "Heading 1" \\s}-{SEQ 图 \\* ARABIC \\s 1} 标题
    """
    all_runs = []

    # "图" / "表" 前缀
    all_runs.append(make_text_run(caption_type))

    # STYLEREF 域: 章节号
    # WHY: 数字 1 引用 Heading 1 样式级别, \\s 返回段落编号
    styleref_runs = make_field_runs(' STYLEREF 1 \\s ')
    all_runs.extend(styleref_runs)

    # 连字符
    all_runs.append(make_text_run('-'))

    # SEQ 域: 章内自动编号
    # WHY: \\s 1 在每个 Heading 1 处重新计数
    seq_runs = make_field_runs(f' SEQ {caption_type} \\* ARABIC \\s 1 ')
    all_runs.extend(seq_runs)

    # 空格
    all_runs.append(make_text_run(' '))

    if rpr_source is not None:
        apply_rpr_to_runs(all_runs, rpr_source)

    return all_runs


def set_caption_format(para):
    """设置题注段落格式：居中、无首行缩进。"""
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


def has_drawing(para_elem):
    """检查段落是否包含图片。"""
    return bool(para_elem.findall('.//' + qn('w:drawing')))


def resize_images(doc, target_width_cm=IMAGE_TARGET_WIDTH_CM):
    """将文档中所有嵌入图片宽度调整为目标宽度，保持纵横比。

    WHY: Pandoc 按原始像素/DPI 嵌入图片，常导致图片过小。
    统一调整为接近页面宽度以确保可读性。
    """
    target_width_emu = Cm(target_width_cm)
    count = 0

    for p in doc.paragraphs:
        inlines = p._element.findall('.//' + qn('wp:inline'))
        for inline in inlines:
            extent = inline.find(qn('wp:extent'))
            if extent is None:
                continue

            cx = int(extent.get('cx', 0))
            cy = int(extent.get('cy', 0))
            if cx == 0 or cy == 0:
                continue

            # 计算缩放比例保持纵横比
            ratio = target_width_emu / cx
            new_cx = target_width_emu
            new_cy = int(cy * ratio)

            extent.set('cx', str(new_cx))
            extent.set('cy', str(new_cy))

            # 同步更新 a:ext（图片实际渲染尺寸）
            for a_ext in inline.findall('.//' + qn('a:ext')):
                a_ext.set('cx', str(new_cx))
                a_ext.set('cy', str(new_cy))

            count += 1

    if count > 0:
        print(f'  {count} image(s) resized to {target_width_cm}cm width')


def process_captions(doc):
    """处理所有图表题注。"""
    count = 0
    paragraphs = doc.paragraphs

    for i in range(len(paragraphs)):
        p = paragraphs[i]
        text = p.text.strip()

        match = CAPTION_PATTERN.match(text)
        if not match:
            continue

        if len(text) > MAX_CAPTION_LENGTH:
            continue

        caption_type = match.group(1)

        # 图片题注：上一段落应含图片
        if caption_type == '图' and i > 0:
            prev_p = paragraphs[i - 1]
            if not has_drawing(prev_p._element) and not has_drawing(p._element):
                continue

        # 提取标题文本
        title_text = CAPTION_PATTERN.sub('', text)

        # 保存格式
        first_run_elem = p.runs[0]._element if p.runs else None

        # 清空 runs
        for r in p._element.findall(qn('w:r')):
            p._element.remove(r)

        # 构建域 runs（全域化：STYLEREF + SEQ）
        caption_runs = build_caption_runs(caption_type, first_run_elem)

        # 标题文本 run
        r_title = make_text_run(title_text)
        if first_run_elem is not None:
            source_rpr = first_run_elem.find(qn('w:rPr'))
            if source_rpr is not None:
                r_title.insert(0, copy.deepcopy(source_rpr))

        # 插入 runs
        ppr = p._element.find(qn('w:pPr'))
        insert_after = ppr if ppr is not None else None

        for run_elem in caption_runs + [r_title]:
            if insert_after is not None:
                insert_after.addnext(run_elem)
                insert_after = run_elem
            else:
                p._element.append(run_elem)

        # 居中、无缩进
        set_caption_format(p)

        count += 1
        print(f'  ok {caption_type}: [{i}] "{caption_type}?-?  {title_text[:30]}"')

    return count


def main():
    if len(sys.argv) < 2:
        print('Usage: python add_captions.py <docx>')
        sys.exit(1)

    docx_path = sys.argv[1]
    print(f'  Processing: {docx_path}')

    doc = Document(docx_path)

    # 处理题注
    count = process_captions(doc)

    if count > 0 or True:  # 即使无题注也保存（可能只调了图片尺寸）
        doc.save(docx_path)
        print(f'  {count} caption(s) processed')
        print(f'  Tip: In Word press Ctrl+A then F9 to update fields')


if __name__ == '__main__':
    main()
