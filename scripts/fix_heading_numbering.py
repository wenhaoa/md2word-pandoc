#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fix_heading_numbering.py - 标题里的手写编号与模板自动编号去重

原则:
    标题的编号由**模板的自动编号**负责，正文文字里不再保留手写编号。
    源 md 若写成 `## 1 概述`、`### 1.1 卫星概况`，转换后 Word 会显示
    「1 1 概述」「1.1 1.1 卫星概况」——本脚本删除文字里的手写编号，
    保留模板自动编号，结果为「1 概述」「1.1 卫星概况」。

做法:
    对每个标题段落，若文本以 `数字[.数字…]` + 分隔符 开头，从 run 序列里
    按字符数删除该前缀（只截断、不重排，跨 run 也安全）。

用法:
    python fix_heading_numbering.py <docx> [--keep-manual]
"""
import re
import sys

from docx import Document

NUM_PREFIX = re.compile(r"^[\s\u3000]*(?!\d{4}(?:\D|$))\d+(?:[.．]\d+)*(?:[\s\u3000]+|[、.．,，][\s\u3000]*)")
HEADING = re.compile(r"^heading\s*(\d+)$")


def heading_level(p):
    try:
        m = HEADING.match(p.style.name.strip().lower())
    except Exception:
        return None
    return int(m.group(1)) if m else None


def strip_prefix(p, n):
    """从段落 run 序列开头删除 n 个字符（只截断，不重排）。"""
    left = n
    for r in p.runs:
        if left <= 0:
            break
        t = r.text or ""
        if not t:
            continue
        cut = min(left, len(t))
        r.text = t[cut:]
        left -= cut
    return n - left


def fix_captions(doc):
    """兼容旧调用；题注由add_captions处理，不再删除标题内的数字。"""
    return 0


def run(path, keep_manual=False):
    doc = Document(path)
    heads = [p for p in doc.paragraphs if heading_level(p)]
    if not heads:
        print("  未发现标题段落，跳过")
        return
    fixed = 0
    for p in ([] if keep_manual else heads):
        txt = p.text or ""
        m = NUM_PREFIX.match(txt)
        if not m:
            continue
        # 只删“纯编号”前缀：避免把 “2026年任务” 这类真内容开头当编号
        if len(m.group(0)) >= len(txt):
            continue
        if strip_prefix(p, m.end()):
            fixed += 1
    # Source caption numbers are parsed by add_captions; never strip a real
    # title beginning with a number (for example, 300km离轨标准).
    cap = 0
    doc.save(path)
    print("  标题 %d 个；删除手写编号 %d 处（保留模板自动编号）" % (len(heads), fixed))
    print("  题注手写序号删除 %d 处" % cap)


def main():
    argv = sys.argv[1:]
    args = [a for a in argv if not a.startswith("--")]
    if not args:
        print(__doc__)
        sys.exit(2)
    run(args[0], '--keep-manual' in argv)


if __name__ == "__main__":
    main()
