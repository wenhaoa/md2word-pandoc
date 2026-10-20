# md2word-pandoc

将Markdown技术报告转换为带封面、目录、可编辑公式、图表题注及交叉引用的Word文档。

## 主要功能

| 功能 | 说明 |
| --- | --- |
| 公式与编号 | 显示公式转OMML，支持源标签、按章编号及正文公式引用 |
| 交叉引用 | 图、表、公式和章节引用转为书签及REF域 |
| 公式排版 | 简单行内公式转文本及上下标，复杂公式保留，合并同属性数学run |
| 表格 | 按内容和正文页宽估算列宽，居中单元格、加粗表头 |
| 题注 | 兼容图片alt题注，支持独立题注、全文连续或章内连续编号 |
| 中文排版 | 保留原有空格和中文引号清理，可用`--preserve-text`跳过 |
| 标题与封面 | 三种标题映射模式，标题回退，封面标题和年月占位符 |
| 文献引用 | 显式使用`--citations`将方括号文献引用整体上标 |
| 预检与校验 | 检查本地资源和悬空引用，校验DOCX图片、数学结构、书签及域 |
| 输出保护 | 独立临时目录、防覆盖输出、并发转换、源文件哈希检查及失败诊断 |

## 安装与使用

需要Pandoc、Node.js、Python及`python-docx`、`docxcompose`、`lxml`。环境检查见[环境配置](references/first_time_setup.md)。

保留原有安装更新入口：

```powershell
powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\.codex\skills\md2word-pandoc\scripts\install_or_update.ps1"
```

无需安装快捷方式即可从skill目录执行：

```powershell
node scripts/check_markdown.js examples/示例技术报告.md --json
node scripts/run_conversion.js examples/示例技术报告.md --citations
```

输出位于源文件同目录，文件名为`<源文件名>YYYYMMDD_HHMMSS.docx`，例如`报告20260507_144822.docx`；仅同秒重名时追加序号以免覆盖。打开Word后更新目录和全部域，再检查分页、表格、图片和公式位置。命令行参数、源文档写法及保护边界见[SKILL.md](SKILL.md)。

## 验证

```powershell
python scripts/test_regressions.py
python scripts/merge_cover_placeholder.test.py
node scripts/run_conversion_filename.test.js
```

本次功能整合日期：2026-10-20。
