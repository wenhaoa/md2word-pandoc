# 技术原理与开发备忘

## 转换流程

```
1    预处理（CJK 空格清理 + 中文引号转换）
2    Pandoc 转换
       --from markdown-smart-subscript-superscript
       --reference-doc=模板.docx
       --lua-filter=style_filter.lua
       --resource-path=源文件目录
2.5  合并封面与目录（merge_cover.py）
2.6  添加图表题注 SEQ 域（add_captions.py）
3    重命名输出文件（北京时间戳）
```

## 核心 Pandoc 命令

```powershell
pandoc "源文件.md" -o "输出.docx" ^
    --from markdown-smart-subscript-superscript ^
    --reference-doc="md2word模板.docx" ^
    --lua-filter="style_filter.lua" ^
    --resource-path="源文件所在目录" ^
    --standalone
```

**参数说明**：
- `-smart`：禁用 smart typography，防止破坏预处理后的中文引号
- `-subscript-superscript`：禁用 `~` 下标和 `^` 上标（中文常用 `~` 表示范围如 300~400km）
- `--resource-path`：确保从任意工作目录调用（如右键发送到）时也能找到相对路径的图片

## 后处理脚本

| 脚本               | 功能                                      | 依赖              |
| ------------------ | ----------------------------------------- | ----------------- |
| `merge_cover.py`   | 合并模板中的封面+目录页，替换 `{{TITLE}}` | python-docx       |
| `add_captions.py`  | 将 `图N-M` / `表N-M` 转为 Word SEQ 域     | python-docx, lxml |
| `style_filter.lua` | 智能标题映射、编号清洗、表格样式          | Pandoc 内置       |

## SEQ 域代码结构

```
图{STYLEREF 1 \s}-{SEQ 图 \* ARABIC \s 1} 标题文本
```
- `STYLEREF 1 \s`：从最近的 Heading 1 获取章节编号
- `SEQ 图 \* ARABIC \s 1`：章内自动递增，每个 Heading 1 处重置
