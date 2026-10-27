#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
autofit_tables.py - 表格列宽自适应 + 表后留白

目标: 以"整表换行行数最少"为准则设定列宽，且**只统计非标题行**（标题行不参与优化）。

算法（等比满足度二分）:
  1 c_i = 第 i 列自然宽度 = 非标题行中该列单行排下的最大宽度
  2 wmin_i = 该列不可断词的最大宽度（CJK 可任意断取 1 字；数字/英文取整串）
  3 求最大 λ 使 Σ clamp(λ·c_i, wmin_i, c_i) ≤ 页宽

例外与补充:
  - **封面表格保持原状**：位于第一个 Heading 1 之前的表格一律不处理
  - 每个正文表格之后插入一个空段落，使表格与后文拉开距离

用法:
    python autofit_tables.py <docx> [--dry-run]
"""
import math
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from table_sections import section_for_block

EMU_PER_PT = 12700
DEFAULT_FONT_PT = 10.5
CELL_PAD_PT = 2 * 5.4
RESERVE_PT = 5.25          # 每列宽度余量：半个中文字符（1 个字符会让多列表格被推出页宽）
ASCII_RATIO = 0.55


# 分字符类别的宽度系数（以字号为 1 em）
# WHY: 统一按 0.55 会把大写字母严重低估——MON_A/MMH_B 这类标识符
# 在 Word 里比估算宽 4~5 pt，导致该列折行。
W_UPPER = 0.68      # 大写字母 M O N A …
W_LOWER = 0.52      # 小写字母
W_DIGIT = 0.55      # 数字
W_PUNCT = 0.50      # 下划线、连字符、括号、斜杠等


def text_width_pt(s, fs=DEFAULT_FONT_PT):
    w = 0.0
    for ch in s:
        if "\u2e80" <= ch <= "\u9fff" or "\uff00" <= ch <= "\uffef":
            w += fs                      # 中日韩全角
        elif "A" <= ch <= "Z":
            w += fs * W_UPPER
        elif "a" <= ch <= "z":
            w += fs * W_LOWER
        elif "0" <= ch <= "9":
            w += fs * W_DIGIT
        else:
            w += fs * W_PUNCT
    return w


def _tokens(s):
    out, cur = [], []
    for ch in s:
        if ch.isspace() or "\u2e80" <= ch <= "\u9fff" or "\uff00" <= ch <= "\uffef":
            if cur:
                out.append("".join(cur))
                cur = []
        else:
            cur.append(ch)
    if cur:
        out.append("".join(cur))
    return out or [s]


def longest_token_pt(s, fs=DEFAULT_FONT_PT):
    return max([text_width_pt(t, fs) for t in _tokens(s)] + [fs])


def col_widths(rows, avail_pt, fs=DEFAULT_FONT_PT, pad_pt=CELL_PAD_PT, reserve_pt=RESERVE_PT):
    """rows 为**非标题行**。返回 (widths, nat, wmin)。"""
    n = max(len(r) for r in rows)
    nat, wmin = [0.0] * n, [0.0] * n
    for r in rows:
        for i, cell in enumerate(r):
            t = (cell or "").strip()
            nat[i] = max(nat[i], text_width_pt(t, fs) + pad_pt)
            wmin[i] = max(wmin[i], longest_token_pt(t, fs) + pad_pt)

    # 估算字宽与 Word 实际字体有偏差：每列至少留 1 个中文字符的余量
    for i in range(n):
        nat[i] += reserve_pt
        wmin[i] += reserve_pt

    if sum(wmin) >= avail_pt:
        k = avail_pt / sum(wmin)
        return [w * k for w in wmin], nat, wmin

    lo, hi = 0.0, 1.0
    for _ in range(60):
        mid = (lo + hi) / 2.0
        tot = sum(min(max(mid * nat[i], wmin[i]), nat[i]) for i in range(n))
        if tot <= avail_pt:
            lo = mid
        else:
            hi = mid
    w = [min(max(lo * nat[i], wmin[i]), nat[i]) for i in range(n)]
    R = avail_pt - sum(w)
    if R > 0.5:
        w = [x + R / n for x in w]
    return w, nat, wmin



def refine_greedy(rows, ws, avail_pt, fs=DEFAULT_FONT_PT):
    """贪心精修：每次把宽度给"每 pt 换行减少最多"的列，直到预算用完。

    WHY: 等比满足度是全局均衡，不保证每张表都达到最小行数；
    贪心按边际收益分配，能进一步压低总行数。
    """
    for _ in range(300):
        if sum(ws) >= avail_pt - 0.5:
            break
        cur = count_lines(rows, ws)
        best = None
        n = len(ws)
        for i in range(n):
            need = sorted(set(
                text_width_pt((r[i] if i < len(r) else ""), fs) + CELL_PAD_PT + RESERVE_PT
                for r in rows))
            for nw in need:
                if nw <= ws[i] + 0.5:
                    continue
                delta = nw - ws[i]
                if sum(ws) + delta > avail_pt:
                    continue
                trial = list(ws)
                trial[i] = nw
                gain = cur - count_lines(rows, trial)
                if gain > 0:
                    score = gain / delta
                    if best is None or score > best[0]:
                        best = (score, i, nw)
                break
        if not best:
            break
        ws[best[1]] = best[2]
    return ws

def table_rows(table):
    return [["".join(p.text for p in c.paragraphs) for c in row.cells]
            for row in table.rows]


def count_lines(rows, ws, fs=DEFAULT_FONT_PT, pad_pt=CELL_PAD_PT):
    """rows 为**非标题行**。"""
    total = 0
    for r in rows:
        mx = 1
        for i, c in enumerate(r):
            if i >= len(ws) or ws[i] <= 0:
                continue
            need = text_width_pt((c or "").strip(), fs) + pad_pt
            mx = max(mx, math.ceil(need / ws[i]))
        total += mx
    return total


def set_widths(table, widths_pt):
    tbl = table._tbl
    tblPr = tbl.tblPr
    for old in tblPr.findall(qn("w:tblLayout")):
        tblPr.remove(old)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)

    for old in tblPr.findall(qn("w:tblW")):
        tblPr.remove(old)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(sum(widths_pt) * EMU_PER_PT / 635)))
    tblPr.append(tblW)

    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement("w:tblGrid")
    for wp in widths_pt:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(wp * EMU_PER_PT / 635)))
        grid.append(gc)
    tblPr.addnext(grid)

    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i >= len(widths_pt):
                continue
            tcPr = cell._tc.get_or_add_tcPr()
            for old in tcPr.findall(qn("w:tcW")):
                tcPr.remove(old)
            tcW = OxmlElement("w:tcW")
            tcW.set(qn("w:type"), "dxa")
            tcW.set(qn("w:w"), str(int(widths_pt[i] * EMU_PER_PT / 635)))
            tcPr.insert(0, tcW)



def set_word_autofit(table):
    """交给 Word 自己排：tblLayout=autofit + 表宽 100% + 删除 tcW。

    WHY: 用估算字宽去模拟排版，估算偏差会让“能放一行的表”被判成折行。
    Word 的“根据内容自动调整”按真实字宽贴紧内容，全部放得下就单行，
    放不下才折行，正好是目标规则。
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    for tag in ("w:tblLayout", "w:tblW"):
        for old in tblPr.findall(qn(tag)):
            tblPr.remove(old)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "autofit")
    tblPr.append(layout)
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:type"), "pct")
    tw.set(qn("w:w"), "5000")
    tblPr.append(tw)
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for old in tcPr.findall(qn("w:tcW")):
                tcPr.remove(old)


def available_width_pt(doc, block=None):
    # WHY: 合并封面后 sections[0] 是封面节（边距 85pt），正文节边距 72pt。
    # 表格在正文里，必须用**最后一节**的正文宽度，否则少算约 26pt，
    # 本可单行放下的窄表会被误判成“需要折行”。
    if block is not None:
        from docx.section import Section
        sec = Section(section_for_block(doc, block), doc.part)
    else:
        sec = doc.sections[-1]
    try:
        return (sec.page_width - sec.left_margin - sec.right_margin) / EMU_PER_PT
    except TypeError:
        return 450.0


def cover_cutoff(doc):
    """第一个 Heading 1 段落在 body 中的位置；其前的表格属于封面/目录。"""
    body = list(doc.element.body)
    for p in doc.paragraphs:
        try:
            if p.style is not None and p.style.name.lower() == "heading 1":
                return body.index(p._element)
        except ValueError:
            continue
    return 0


def run(path, dry=False, native=False, body_only=False, table_indices=None, compact=False):
    doc = Document(path)
    avail = available_width_pt(doc)
    body = list(doc.element.body)
    cut = 0 if body_only else cover_cutoff(doc)
    print("  可用正文宽度 %.0f pt；封面区元素 %d 个" % (avail, cut))

    tables = list(doc.tables)
    if native:
        n = 0
        for table_index, tb in enumerate(tables):
            if table_indices is not None and table_index not in table_indices:
                continue
            try:
                idx = body.index(tb._tbl)
            except ValueError:
                continue
            if idx < cut:
                continue
            set_word_autofit(tb)
            n += 1
        if not dry:
            doc.save(path)
        print("  已改为 Word 自动调整表格宽度的表：%d 张（跳过封面 %d 张）"
              % (n, sum(1 for tb in tables
                        if body.index(tb._tbl) < cut if tb._tbl in body)))
        return

    skipped = 0
    before = after = 0
    processed = []

    for table_index, t in enumerate(tables):
        if table_indices is not None and table_index not in table_indices:
            continue
        try:
            idx = body.index(t._tbl)
        except ValueError:
            idx = 10 ** 9
        if idx < cut:                      # 封面表格：保持原状
            skipped += 1
            continue
        avail = available_width_pt(doc, t._tbl)
        rows = table_rows(t)
        if len(rows) < 2:
            continue
        header, data = rows[0], rows[1:]
        font_pt = DEFAULT_FONT_PT
        pad_pt = 4.0 if compact else CELL_PAD_PT
        reserve_pt = 1.0 if compact else RESERVE_PT
        if compact:
            # Keep numeric times together before shrinking below ordinary small-five type.
            while font_pt > 9.0:
                _, _, minimum = col_widths(data, avail, font_pt, pad_pt, reserve_pt)
                if sum(minimum) <= avail:
                    break
                font_pt -= 0.5
            if not dry:
                from docx.shared import Pt
                for row in t.rows:
                    for cell in row.cells:
                        props = cell._tc.get_or_add_tcPr()
                        margins = props.find(qn('w:tcMar'))
                        if margins is None:
                            margins = OxmlElement('w:tcMar')
                            props.append(margins)
                        for side in ('left', 'right'):
                            item = margins.find(qn('w:' + side))
                            if item is None:
                                item = OxmlElement('w:' + side)
                                margins.append(item)
                            item.set(qn('w:w'), '40')
                            item.set(qn('w:type'), 'dxa')
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(font_pt)

        # 规则：若**所有行（含标题行）**都能单行放下 → 直接用最小列宽
        ncol = max(len(r) for r in rows)
        # 估算字宽与 Word 实际字体有偏差，每列多留 1 个中文字符的余量
        RESERVE = reserve_pt

        def nat_of(rs):
            out = []
            for col in range(ncol):
                m_ = 0.0
                for r in rs:
                    cell = (r[col] if col < len(r) else "") or ""
                    m_ = max(m_, text_width_pt(cell, font_pt) + pad_pt)
                out.append(m_ + RESERVE)
            return out

        # 三级判定：
        # ① 所有单元格（含标题行）单行最大列宽+余量，总宽不超页宽 -> 采用
        # ② 第2行至结尾（数据行）单行最大列宽+余量，总宽不超页宽 -> 采用
        # ③ 否则对数据行做最优分配（整表不超页宽、换行数最小）
        nat_all = nat_of(rows)
        nat_data = nat_of(data)
        if sum(nat_all) <= avail:
            widths, nat, wmin = nat_all, nat_all, nat_all
            branch = "1-全行单行"
        elif sum(nat_data) <= avail:
            widths, nat, wmin = nat_data, nat_data, nat_data
            branch = "2-数据行单行"
        else:
            widths, nat, wmin = col_widths(data, avail, font_pt, pad_pt, reserve_pt)
            branch = "3-最优分配"
        tot = sum(widths)
        if tot > avail:
            widths = [w * avail / tot for w in widths]

        print("     表%d: 全行%.0f / 数据行%.0f / 可用%.0f -> %s"
              % (len(processed) + 1, sum(nat_all), sum(nat_data), avail, branch))

        orig = []
        grid = t._tbl.find(qn("w:tblGrid"))
        if grid is not None:
            for gc in grid.findall(qn("w:gridCol")):
                orig.append(int(gc.get(qn("w:w")) or 0) * 635 / EMU_PER_PT)
        if sum(widths) > avail + 0.5:
            print("     !! 表%d 计算总宽 %.0f > 可用 %.0f（列数 %d，rows %d）"
                  % (len(processed) + 1, sum(widths), avail, len(widths), len(rows)))
        lb = count_lines(data, orig, font_pt, pad_pt) if len(orig) == len(widths) else count_lines(data, nat, font_pt, pad_pt)
        la = count_lines(data, widths, font_pt, pad_pt)
        before += lb
        after += la
        if not dry:
            set_widths(t, widths)
        processed.append((len(processed) + 1, widths, lb, la))
        if compact:
            print('     横向表紧凑排版：字号%.1f pt，单侧留白2 pt' % font_pt)

    if not dry:
        for table_index, t in enumerate(tables):                   # 表后补空段
            if table_indices is not None and table_index not in table_indices:
                continue
            try:
                idx = body.index(t._tbl)
            except ValueError:
                continue
            if idx < cut:
                continue
            nxt = t._tbl.getnext()
            if nxt is not None and nxt.tag == qn("w:p"):
                txt = "".join(x.text or "" for x in nxt.findall(".//" + qn("w:t")))
                if not txt.strip():
                    continue               # 已有空段
            t._tbl.addnext(OxmlElement("w:p"))
        doc.save(path)

    for i, w, lb, la in processed[:6]:
        print("  表%-3d 列宽(pt) %-44s 非标题行 %d→%d"
              % (i, " ".join("%.0f" % x for x in w), lb, la))
    print("  处理表格 %d 张（跳过封面 %d 张）" % (len(processed), skipped))
    print("  非标题行换行合计 %d → %d（%s）"
          % (before, after, "dry-run" if dry else "已保存"))


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        sys.exit(2)
    run(args[0], "--dry-run" in sys.argv, "--word-autofit" in sys.argv,
        "--body-only" in sys.argv)


if __name__ == "__main__":
    main()
