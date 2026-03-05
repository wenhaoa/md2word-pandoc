# 📄 md2word-pandoc

> **Markdown → Word 高精度转换** ｜ Antigravity AI Skill

[![Pandoc](https://img.shields.io/badge/Pandoc-2.x+-4B8BBE)](https://pandoc.org)
[![Node.js](https://img.shields.io/badge/Node.js-14+-339933)](https://nodejs.org)
[![Platform](https://img.shields.io/badge/Platform-Windows-blue)]()

将 Markdown 文档高精度转换为排版精美的 Word 文件，专为**中文技术报告**场景优化。

## ✨ 核心能力

| 能力           | 说明                                   |
| -------------- | -------------------------------------- |
| 🔢 原生公式     | LaTeX → Word OMML 原生公式，完美可编辑 |
| 📑 智能标题映射 | 自动检测 H1 数量，适配不同写作习惯     |
| 🎨 模板样式定制 | 直接在 Word 模板中修改样式即生效       |
| 🇨🇳 中文排版优化 | 自动修正引号方向、清理 CJK/Latin 空格  |
| 📋 封面与目录   | 自动合并封面，`{{TITLE}}` 占位符注入   |
| 🏷️ 图表自动编号 | `图N-M` / `表N-M` 转为 Word SEQ 域     |

## 🚀 安装

```powershell
# 通过 AI 一键安装
npx @anthropic/skills-cli add https://github.com/wenhaoa/md2word-pandoc
```

**环境依赖**：Pandoc 2.x+、Node.js 14+、Python 3.x + python-docx

首次使用时 AI 会自动检查并配置配套的写作规则和预检工作流，详见 [first_time_setup.md](references/first_time_setup.md)。

## 📖 使用

- **让 AI 帮转**：对 AI 说"帮我把 xxx.md 转成 Word"
- **双击/拖拽/右键发送到**：首次运行 `scripts\install_shortcuts.bat` 安装快捷方式
- **命令行**：`node scripts\run_conversion.js "文件.md"`

技术细节、样式定制、常见问题等完整文档见 [SKILL.md](SKILL.md)。

## 📁 目录结构

```
md2word-pandoc/
├── SKILL.md                    # AI 核心参考（技术细节+维护指南）
├── README.md                   # 项目简介（本文件）
├── scripts/
│   ├── run_conversion.js       # 主转换流程
│   ├── style_filter.lua        # 智能标题过滤器
│   ├── add_captions.py         # 图表题注 SEQ 域处理
│   ├── merge_cover.py          # 封面合并
│   ├── md2word_gui.bat         # GUI 入口（双击/拖拽/SendTo）
│   └── install_shortcuts.bat   # 快捷方式安装
├── templates/
│   └── md2word模板.docx        # Word 样式模板
├── references/
│   ├── first_time_setup.md     # 首次安装指南（含 Rules + Workflow）
│   └── technical_details.md    # 转换流程技术备忘
└── examples/
    └── 示例技术报告.md          # 示例文档
```

## 📜 版本

当前版本 **V1.3** (2026-03-04)。完整版本历史见 [SKILL.md](SKILL.md#版本历史)。
