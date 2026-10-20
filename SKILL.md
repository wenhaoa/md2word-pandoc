---
name: md2word-pandoc
description: >-
  Convert Markdown technical reports to Word/docx with Pandoc, a Chinese Word
  template, editable OMML equations, captions and cross-references. Use for
  转Word、转docx、md转word、markdown转docx and Markdown preflight before export.
---

# Markdown转Word

## 依赖与入口

需要Pandoc、Node.js、Python3及`python-docx`、`docxcompose`、`lxml`。模板为`templates/md2word模板.docx`，过滤器为`scripts/style_filter.lua`。首次使用或转换失败时检查依赖；缺少Python包时用同一个Python解释器安装。

```powershell
pandoc --version
node --version
python --version
python -m pip show python-docx docxcompose lxml
```

从skill根目录执行：

```powershell
node scripts/check_markdown.js 文件.md --json
node scripts/run_conversion.js 文件.md
```

从任意目录执行时使用脚本绝对路径，Codex默认安装位置为`~/.codex/skills/md2word-pandoc`。`MD2WORD_SKILL_DIR`可指定另一份完整skill；脚本、过滤器和模板必须来自同一目录。

保留原有安装更新、PowerShell函数、双击、拖拽和右键“发送到”入口。安装更新使用`scripts/install_or_update.ps1`，仅转换时不运行安装脚本。依赖和快捷入口说明见[环境配置](references/first_time_setup.md)。

## 内容保护

转换不回写Markdown源文件。默认恢复原有预处理：清理中英文及数字单位间的多余空格、转换中文双引号、连接中文软换行；只处理本次临时输入，保护YAML、代码、链接、路径和数学表达式。`--preserve-text`可跳过这层排版清理。正文改写或文字润色仍需单独处理并复核差异。

Pandoc先解析Markdown为AST，只在显示公式节点中提取并去掉`\tag`。通过参数数组调用程序，标题和路径不拼接为shell代码。每次创建独立临时目录，输出恢复为`<源文件名>YYYYMMDD_HHMMSS.docx`并独占写入；仅同秒重名时追加`_2`、`_3`等序号。发布前检查源文件哈希；转换期间源文件变化时终止输出。

成功后仅删除本次创建的临时目录；必要步骤失败返回非零并保留诊断目录，不宣布成功。不得清理源文件目录中既有的`temp_input.md`、`temp_output.docx`或其他文件。

## 源文档与题注

新写的中文报告遵守用户的Markdown转Word规范：用YAML的`title`设置封面标题，正文从`## 1. 概述`开始。保留已有合理层级和有意义的引用，不为转换擅自改写技术内容。封面标题依次取YAML的`title`、唯一总标题或文件名。

兼容原有`![图1-1 标题](path)`写法，图片alt自动生成可见题注；也支持图下另写独立题注。已有独立题注时优先使用它，不重复生成。图题在图下，表题在表上，题注与相邻内容之间保留一个空行。

```markdown
---
title: 技术报告
---

## 1. 概述

![图1-1 系统架构](images/system.png)

图1-1 系统架构

表1-1 输入参数

| 参数 | 取值 | 说明 |
| -------- | ------------------ | -------------------------------------------- |
| 推力 | 20mN | 设计输入 |

见图1-1和表1-1。
```

支持`图1 标题`、`图4-1 标题`、旧式`图4.1 标题`、普通或加粗题注及全角空格。点号分隔的旧式题注统一输出为章内连字符编号。沿用全文连续或章内连续编号，每组从1开始且不重复。脚本题注识别上限80字符，新写报告仍按用户要求控制在60字符以内。无编号短题注仅在明确紧邻图表时识别；含复杂内联对象的题注保留原文，不强行重建。同名题注保留，重复编号报错。

图表题注转为SEQ域，章内编号同时使用STYLEREF；缓存保留源编号，不强制标记为dirty，避免Word打开时弹出域更新确认框；手动F9仍可更新。表题段前、图题段后各留0.5行。图片居中；表格单元格居中并清除缩进，表头加粗。表格`<br>`转换为Word换行。

## 标题、封面和表格

默认`--heading-mode auto`：无H1时H2映射为Heading1；唯一未带章节编号的H1作为封面总标题；多个H1或明确带编号的单个H1作为正文章节。需要明确指定时使用`--heading-mode shift`或`--heading-mode chapters`。

手写章节编号只在Lua过滤器中删除一次，增加“一、”“五．”等中文序号识别，保留“2026年任务”“3D模型”“3.14km测量”等实际标题文字。`fix_heading_numbering.py`是兼容工具，主流程不重复调用。模板应为Heading1至Heading6配置一致的多级编号；附录字母及特殊起始章号需另行确认。

保留现有封面模板和`{{TITLE}}`、`{{DATE_CN}}`占位符，并兼容模板中的字面年月。默认按北京时间更新年月；`merge_cover.py --date-cn`可指定年月。合并时保持嵌套书签的起止配对，仅删除紧接分节符的结构性空段，不删除图片、公式、域或分页内容。

表格列宽采用三级判定：能容纳全部单元格时保持全表单行；仅数据行能容纳时允许表头折行；否则按内容比例分配并限制总宽不超正文页宽。字宽按中英文、数字及标点估算，列间保留余量。正文表格在合并封面前处理，封面表格不参与调整。新写Markdown表格仍按用户要求设置不等长的分隔线。

## 公式、交叉引用和参考文献

显示公式保留可编辑OMML，支持`$$...$$`或`\[...\]`，编号使用`\tag{2-1}`，正文按`式(2-1)`引用。缺少标签时按章和顺序补号；已给出的标签必须与章节及顺序一致，否则终止，不静默修改源引用。`equation_numbering.py --plan`和`--plan-file`读取源编号，重复执行不会追加第二份编号。

公式段落左对齐、无缩进，通过中心和右侧制表位定位公式及编号；行距1.5倍、段前段后各0.5行。简单、无特殊样式的行内公式可转为普通文本及上下标；包含分式、根式、矩阵、嵌套脚标等结构时保留整个OMML。只合并属性一致的相邻数学run，将公式内U+2000–U+200A特殊空格归一化；不执行正文数字单位空格删除。

图、表、公式和章节按实际编号建立书签并生成REF域，如`图2-1`、`表2-1`、`式(2-1)`、`第4章`、`第4.8节`。普通文本run及表格单元格内的完整引用参与处理，保留原格式；已有域、代码和跨run引用不强行重建，跨run引用需人工检查。

参考文献上标默认关闭。仅在确认正文方括号数字均为文献引用时加`--citations`，将`[1]`、`[1,2]`、`[1-3]`整体上标，保留参考文献章节及条目编号。

## 执行与验证

1. 定位源文件，检查依赖，运行`check_markdown.js`。新增检查包括本地图片和链接、图表与公式编号、悬空引用；代码示例不作为正文处理。`--json`输出结构化结果。预检与转换是两个独立命令，主转换脚本不自动运行预检。
2. 按用户已授权的范围修复源文件。用户只要求转换时，不擅自改写章节或技术内容。
3. 运行`run_conversion.js`。按需使用`--heading-mode auto|shift|chapters`、`--citations`、`--no-caption`、`--preserve-text`、`--open`。`--no-caption`跳过题注及图表引用域化，公式和章节引用仍处理。`--open`等待系统启动命令完成并报告打开错误，不等待Word关闭；右键“发送到”、双击和拖拽入口继续默认自动打开。
4. 主流程记录Pandoc正文基线，再处理行内公式、数学run、题注、表格、公式编号和引用，完成正文结构校验后合并封面并再次校验。
5. 校验ZIP、图片哈希、图片对象、表格、显示公式、复杂数学结构、书签起止及REF目标，检查封面占位符残留及编号书签跨段；全部通过且源文件未变化才输出源文件同目录下的独立docx。
6. 报告实际输出路径和验证结果。结构校验不能代替实际分页检查；打开Word更新目录和全部域，核对图片、题注、表格换行、公式位置及分页。未渲染或目视核验时明确说明，不宣称版式验收完成。MathType不是必要依赖。

回归测试：

```powershell
python scripts/test_regressions.py
python scripts/merge_cover_placeholder.test.py
node scripts/run_conversion_filename.test.js
```

测试在系统临时目录运行，覆盖完整转换、并发、内容保留及失败停止。`update_skill_doc.py`和`patch_caption_pipeline.py`为已停用的一次性迁移脚本，不用于更新安装文件。

主要功能示例及验证命令：

```powershell
node scripts/check_markdown.js examples/示例技术报告.md --json
node scripts/run_conversion.js examples/示例技术报告.md --citations
```

示例覆盖图表与公式编号、交叉引用、简单及复杂行内公式、不同内容长度的表格和文献上标。标题回退、并发和失败处理由回归测试覆盖。流程实现见[技术细节](references/technical_details.md)，文字审读提示见[写作检查](references/ai-writing-checklist.md)。

与更新前的行为差异及兼容性核查见[兼容性说明](references/compatibility_review.md)。
