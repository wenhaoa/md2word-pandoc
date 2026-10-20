#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
merge_math_runs.py - 合并公式内相邻的 m:r，消除字符间视觉间隙

问题:
    pandoc 把 LaTeX 公式转 OMML 时，**每个字符输出成一个独立的 m:r**。
    Word 渲染相邻斜体 run 时会各自应用字距，看起来就是字母之间被插了空格，
    例如 m_0=m_static+Δm_trunc 显示成 "m 0 = m static + Δ m trunc"。

做法:
    递归遍历每个 m:oMath，把**相邻且 m:rPr 相同**的 m:r 合并成一个，
    文本拼接进第一个 run 的 m:t（加 xml:space="preserve"）。

效果:
    一段连续符号变成一个 run，Word 按正常字距排版，间隙消失；公式语义不变。

用法:
    python merge_math_runs.py <docx> [--dry-run]
"""
import sys

from docx import Document
from lxml import etree
from docx.oxml.ns import qn

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML = "http://www.w3.org/XML/1998/namespace"


def m(tag):
    return "{%s}%s" % (M, tag)


def rpr_key(run):
    # Both math and Word run properties determine symbol appearance.
    return tuple(etree.tostring(child) for child in run if child.tag != m('t'))


def mtext(run):
    return "".join(t.text or "" for t in run.findall(m("t")))


def merge_in(parent):
    """合并 parent 下的相邻 m:r。返回合并掉的 run 数。"""
    removed = 0
    children = list(parent)
    i = 0
    while i < len(children):
        ch = children[i]
        if ch.tag == m("r"):
            j = i + 1
            key = rpr_key(ch)
            while j < len(children) and children[j].tag == m("r") and rpr_key(children[j]) == key:
                txt = mtext(children[j])
                ts = children[j].findall(m("t"))
                if ts:
                    last = ch.findall(m("t"))[-1] if ch.findall(m("t")) else None
                    if last is not None:
                        last.text = (last.text or "") + txt
                        last.set("{%s}space" % XML, "preserve")
                    else:
                        t = etree.SubElement(ch, m("t"))
                        t.text = txt
                        t.set("{%s}space" % XML, "preserve")
                parent.remove(children[j])
                children = list(parent)
                removed += 1
            i += 1
            continue
        removed += merge_in(ch)
        i += 1
    return removed


# 字体缺字形的特殊空格：U+2001(EM QUAD) 等会显示成方块（tofu）
BAD_SPACES = "\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a"


def normalize_spaces(doc):
    """把公式内的特殊空格换成普通空格，避免缺字形显示为方块。

    WHY: pandoc 把 \\qquad 转成 U+2001(EM QUAD)、\\, 转成 U+2009(THIN SPACE)。
    中文字体常有这些码位缺字形，Word 里就渲染成一个方框。
    """
    n = 0
    # Limit normalization to math text; ordinary text and code remain intact.
    targets = list(doc.element.body.iter(m("t")))
    for t in targets:
        s = t.text or ""
        fixed = s
        if any(ch in BAD_SPACES for ch in s):
            fixed = "".join(" " if ch in BAD_SPACES else ch for ch in s)
            n += sum(1 for ch in s if ch in BAD_SPACES)
        if fixed != s:
            t.text = fixed
            t.set("{%s}space" % XML, "preserve")
    return n


def run(path, dry=False):
    doc = Document(path)
    omaths = list(doc.element.body.iter(m("oMath")))
    total = 0
    for om in omaths:
        total += merge_in(om)
    sp = normalize_spaces(doc)
    if not dry:
        doc.save(path)
    print("  合并相邻 m:r %d 个；修正特殊空格 %d 处（公式 %d 个）%s"
          % (total, sp, len(omaths), "（dry-run）" if dry else ""))


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        sys.exit(2)
    run(args[0], "--dry-run" in sys.argv)


if __name__ == "__main__":
    main()
