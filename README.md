# 📄 md2word-pandoc

> **Markdown → Word 高精度转换**｜Antigravity / Jules AI Skill

[![Pandoc](https://img.shields.io/badge/Pandoc-2.x+-4B8BBE)](https://pandoc.org)
[![Node.js](https://img.shields.io/badge/Node.js-14+-339933)](https://nodejs.org)
[![Platform](https://img.shields.io/badge/Platform-Windows-blue)]()

将 Markdown 文档高精度转换为排版精美的 Word 文件，专为**中文技术报告**场景优化。

---

## ✨ 核心能力

| 能力               | 说明                                                       |
| ------------------ | ---------------------------------------------------------- |
| 🔢 **原生公式**     | LaTeX `$...$` / `$$...$$` → Word OMML 原生公式，完美可编辑 |
| 📑 **智能标题映射** | 自动检测 H1 数量，匹配不同写作习惯 → Word 多级标题编号     |
| 🎨 **模板样式定制** | 直接在 Word 模板中修改字体、编号、间距，即改即生效         |
| 🇨🇳 **中文排版优化** | 自动修正双引号方向、清理 CJK/Latin 多余空格                |
| 📋 **封面与目录**   | 自动合并封面模板，`{{TITLE}}` 占位符注入文档标题           |
| 📊 **表格后处理**   | 自动添加框线、设置居中、消除继承缩进                       |
| 📁 **中文路径兼容** | 自动备份到临时 ASCII 路径，避免 Pandoc 中文路径问题        |

---

## 🚀 一键安装

### 给 AI 发送安装指令

在 Antigravity  对话中告诉 AI：

```
帮我安装这个 skill：https://github.com/wenhaoa/md2word-pandoc
```

AI 会自动执行：
```powershell
npx @anthropic/skills-cli add https://github.com/wenhaoa/md2word-pandoc
```

安装目标路径：`~/.gemini/antigravity/skills/md2word-pandoc/`

### 环境依赖

安装前请确保已有以下软件：

```powershell
# 安装 Pandoc（必须）
winget install JohnMacFarlane.Pandoc

# 安装 Node.js（必须）
winget install OpenJS.NodeJS.LTS

# 安装 Python（封面合并需要）
winget install Python.Python.3.12
pip install python-docx
```

### 首次使用配套

安装完成后，**首次让 AI 执行"转 Word"时**，AI 会自动检查并提示安装配套的写作规则和预检工作流。详见 [references/first_time_setup.md](references/first_time_setup.md)。

---

## 📖 使用方式

### 方式 1：让 AI 帮你转

直接对 AI 说"帮我把 xxx.md 转成 Word"或"转 Word"即可。AI 会自动：
1. 执行格式预检（`/report-check`）
2. 调用转换脚本生成 Word 文件

### 方式 2：命令行

```powershell
# 直接调用
node "$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\scripts\run_conversion.js" "你的文件.md"

# 或配置快捷命令后
md2word "你的文件.md"
```

### 输出

文件生成在**源文件同目录**，格式：`<文件名>_2026-03-03T13-14-37.docx`

---

## 📁 目录结构

```
md2word-pandoc/
├── SKILL.md                          # AI 主参考文档（技术细节）
├── README.md                         # 本文件（GitHub 展示）
├── SETUP_GUIDE.md                    # PowerShell 快捷命令配置
├── scripts/
│   ├── run_conversion.js             # 主转换脚本（预处理+调用 Pandoc）
│   ├── style_filter.lua              # 智能标题过滤器
│   ├── merge_cover.py                # 封面合并脚本
│   └── md2word-function.ps1          # PowerShell 快捷函数
├── templates/
│   └── md2word模板.docx              # Word 样式模板（可自定义）
├── references/
│   ├── first_time_setup.md           # ⭐ 首次安装指南（含配套 Rules + Workflow）
│   └── technical_details.md          # 技术实现细节
└── workflows/
    └── md2word.md                    # /md2word 工作流（供 AI 或用户调用）
```

---

## 📝 Markdown 写作要求

转 Word 的 Markdown 文件需遵循以下基本规范（完整规范见 [first_time_setup.md](references/first_time_setup.md)）：

### 必须：YAML frontmatter 标题

```markdown
---
title: 你的文档标题
---

## 1. 第一章
...
```

`title` 字段会被自动注入到 Word 封面页的 `{{TITLE}}` 占位符中。**不写 title 则封面标题为空**。

### 关键格式规则

- 用 `##` 起步（不用 `#`），过滤器自动映射为 Word Heading 1
- 标题编号带点号：`## 1. 概述`（不是 `## 1 概述`）
- 避免不必要的加粗
- 段落优先，减少分点列举
- 数学公式用 `$...$` 和 `$$...$$`，简单数字+单位直写（如 1100km）

---

## 🔧 自定义模板

1. 打开 `templates/md2word模板.docx`
2. 在 Word 中修改样式（右键标题/正文 → 修改样式）
3. 保存后重新转换即可生效

可自定义项：标题编号格式、正文字体/行距/缩进、表格样式。

---

## ❓ 常见问题

| 问题                  | 解决方案                                          |
| --------------------- | ------------------------------------------------- |
| 表格无框线            | 修改 Word 模板中 Table Normal 样式，添加全框线    |
| 封面标题为空          | 在 md 文件头部添加 `title: 标题` YAML frontmatter |
| 公式字体不一致        | 全选文档设置字体 Cambria Math                     |
| CJK 与 Latin 间距异常 | 已自动处理，如仍有问题检查源文件                  |

更多问题见 [SKILL.md](SKILL.md) 常见问题章节。

---

## 📜 版本历史

| 版本 | 日期       | 更新内容                                                          |
| ---- | ---------- | ----------------------------------------------------------------- |
| V1.2 | 2026-03-03 | 新增首次安装引导、配套 Rules/Workflow 打包、report-check 预检串联 |
| V1.1 | 2026-03-02 | 修复封面标题重复、中文双引号方向、新增表格后处理                  |
| V1.0 | 2026-02-05 | 初始版本                                                          |
