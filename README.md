# md2word-pandoc Skill

[![GitHub](https://img.shields.io/badge/GitHub-wenhaoa%2Fmd2word--pandoc-blue)](https://github.com/wenhaoa/md2word-pandoc)

本 Skill 提供基于 Pandoc 的 Markdown 到 Word 高精度转换方案。

## 快速使用

1. **阅读主文档**：查看 [`SKILL.md`](./SKILL.md) 了解详细技术说明
2. **复制必需文件**到项目目录：
   ```powershell
   # 从 scripts/ 复制
   run_conversion.js
   style_filter.lua
   
   # 从 templates/ 复制
   md2word模板.docx
   ```
3. **修改配置**：编辑 `run_conversion.js` 设置源文件名
4. **执行转换**：`node run_conversion.js`

## 核心特性

- ✅ LaTeX 公式自动转为 Word 原生公式
- ✅ 智能标题映射（自动检测 H1 数量）
- ✅ 所见即所得样式定制
- ✅ 中文排版优化（自动转换双引号、清理 CJK/Latin 间距）
- ✅ 封面与目录合并（`{{TITLE}}` 占位符自动替换）
- ✅ 表格框线与居中自动后处理
- ✅ 自动处理中文文件名和时间戳

## 目录结构

```
md2word-pandoc/
├── SKILL.md              # 完整技术文档
├── README.md             # 本文件
├── scripts/              # 核心脚本
│   ├── run_conversion.js # 主转换脚本（含预处理）
│   ├── style_filter.lua  # Lua 过滤器
│   └── merge_cover.py    # 封面合并脚本
└── templates/            # 样式模板
    └── md2word模板.docx  # Word 样式模板
```

## 相关 Workflow

- `/md2word`：快速执行转换流程（全局可用）
  - 位置：`$env:USERPROFILE\.gemini\antigravity\workflows\md2word.md`

## 问题反馈

遇到问题请参考 `SKILL.md` 中的"常见问题"章节。
