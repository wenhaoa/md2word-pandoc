---
description: Codex 使用的 Markdown 转 Word 预检与转换流程
---

# Markdown 到 Word 转换流程（Codex）

## 1. 触发条件

用户要求将 Markdown、md、报告、论文或技术文档转换为 Word/docx 时，使用本流程。常见表达包括“转 Word”“转 docx”“md 转 word”“把这个 Markdown 生成 Word”。

## 2. 输入定位

1. 如果用户给出 `.md` 路径，使用该文件。
2. 如果用户只说“当前报告”或“这个文档”，扫描当前工作目录下的 `.md` 文件。
3. 如果候选文件多于一个且无法从文件名判断，先询问用户。

## 3. 依赖检查

首次使用或转换失败时运行：

```powershell
pandoc --version
node -v
python --version
python -m pip show python-docx
```

缺少依赖时先报告缺少项和安装命令，不继续转换。

## 4. 自动预检

优先运行预检脚本：

```powershell
node "$env:USERPROFILE\.codex\skills\md2word-pandoc\scripts\check_markdown.js" "源文件.md"
```

脚本输出为带行号的问题表。脚本无法判断的语义问题，再按下表人工审读。

转换前检查以下项目，并按风险处理：

| 检查项 | 可直接修复 | 需要确认 |
| ------ | ---------- | -------- |
| YAML frontmatter 缺少 `title` | 否 | 是 |
| frontmatter 外存在 `---` 水平线 | 是 | 否 |
| 标题格式不是 `## N.` / `### N.N` / `#### N.N.N` | 视情况 | 大范围调整需确认 |
| 标题前后缺少空行 | 是 | 否 |
| 连续空行、行尾空格、空白行内空格 | 是 | 否 |
| 正文手动硬换行 | 视情况 | 合并正文需确认 |
| 图片题注不是 `![图N-M 标题](path)` | 否 | 是 |
| 表格上方缺少 `表N-M 标题` | 否 | 是 |
| 表格题注与表格之间不是一个空行 | 是 | 否 |
| 表格分隔横线过短，列宽未控制 | 是 | 否 |
| 图号/表号章内不连续 | 否 | 是 |
| 正文存在固定图表编号或交叉引用 | 视情况 | 改写内容需确认 |
| 图片行后缺少空行 | 是 | 否 |
| 表格前后缺少空行 | 是 | 否 |
| GitHub 提示块 `> [!NOTE]` | 是 | 否 |
| Mermaid 代码块未转 PNG | 否 | 是 |
| 简单数字+单位被公式包裹 | 是 | 否 |
| 正式报告中存在 `-` 无序列表 | 视情况 | 改写内容需确认 |

预检报告使用表格：

| 行号 | 问题类型 | 当前内容 | 建议修改 |
| ---- | -------- | -------- | -------- |

## 5. 执行转换

默认命令：

```powershell
node "$env:USERPROFILE\.codex\skills\md2word-pandoc\scripts\run_conversion.js" "源文件.md"
```

可选参数：

```powershell
node "$env:USERPROFILE\.codex\skills\md2word-pandoc\scripts\run_conversion.js" "源文件.md" --no-caption
```

## 6. 输出确认

确认源文件同目录生成 `.docx`：

```text
<源文件名>_YYYY-MM-DDTHH-MM-SS.docx
```

最终回复包含：

1. 输出文件路径。
2. 已自动修复的问题。
3. 仍需用户确认或在 Word 中检查的问题。
4. Word 打开后执行 `Ctrl+A` → `F9` 更新图表编号域。

## 7. 失败处理

1. “模板文件不存在”：检查 skill 是否位于 `~/.codex/skills/md2word-pandoc`，或设置 `MD2WORD_SKILL_DIR`。
2. “pandoc 不是内部或外部命令”：安装 Pandoc 并重启终端。
3. Python 脚本失败：先检查 `python-docx`，再尝试加 `--no-caption` 定位是否为题注后处理问题。
4. 图片丢失：确认图片路径相对源 Markdown 所在目录可解析。
