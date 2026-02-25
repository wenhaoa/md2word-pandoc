"""
merge_cover.py - 将封面+目录前缀模板与 Pandoc 生成的正文合并

用法:
    python merge_cover.py <模板docx> <正文docx> <输出docx> [--title "标题"]

依赖:
    pip install docxcompose
"""

import sys
import argparse
from docxcompose.composer import Composer
from docx import Document


def replace_title_placeholder(doc, title):
    """替换模板中的 {{TITLE}} 占位符为实际标题。
    
    WHY: Word 可能将 {{TITLE}} 拆分到多个 run 中，
    因此先尝试 run 级替换，若失败则回退到段落级重组。
    """
    PLACEHOLDER = '{{TITLE}}'
    
    for para in doc.paragraphs:
        if PLACEHOLDER not in para.text:
            continue
        
        # 尝试 run 级精确替换
        replaced = False
        for run in para.runs:
            if PLACEHOLDER in run.text:
                run.text = run.text.replace(PLACEHOLDER, title)
                replaced = True
        
        if replaced:
            continue
        
        # 回退：占位符被拆分到多个 run，用段落级重组
        # WHY: 保留第一个 run 的格式，清空其余 run
        new_text = para.text.replace(PLACEHOLDER, title)
        if para.runs:
            para.runs[0].text = new_text
            for run in para.runs[1:]:
                run.text = ''


def merge(prefix_path, body_path, output_path, title=None):
    """合并前缀模板与正文文档。
    
    Args:
        prefix_path: 包含封面+目录的模板文件路径
        body_path: Pandoc 生成的正文文件路径
        output_path: 合并后的输出文件路径
        title: 可选，替换 {{TITLE}} 占位符的标题文本
    """
    # 1. 打开前缀模板
    prefix = Document(prefix_path)
    
    # 2. 替换标题占位符
    if title:
        replace_title_placeholder(prefix, title)
    
    # 3. 合并正文
    composer = Composer(prefix)
    body = Document(body_path)
    composer.append(body)
    
    # 4. 保存
    composer.save(output_path)
    print(f"   合并完成: {output_path}")


def main():
    parser = argparse.ArgumentParser(description='合并封面模板与正文文档')
    parser.add_argument('prefix', help='封面+目录模板文件路径')
    parser.add_argument('body', help='正文文档路径')
    parser.add_argument('output', help='输出文件路径')
    parser.add_argument('--title', help='替换 {{TITLE}} 占位符的标题', default=None)
    
    args = parser.parse_args()
    merge(args.prefix, args.body, args.output, args.title)


if __name__ == '__main__':
    main()
