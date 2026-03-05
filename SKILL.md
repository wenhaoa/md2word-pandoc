---
name: md2word-pandoc
description: |
  使用 Pandoc 将 Markdown 高精度转换为 Word 文档，支持原生公式、自定义样式和智能标题映射。
  当用户要求将 md 转 word、markdown 转 docx、md 转 docx、把 markdown 文件转换成 word、
  md 文档生成 word 格式时，必须使用本 Skill。即使用户只是简单说"转成 word"或"要 docx"，
  只要涉及 Markdown 源文件，都应该激活本 Skill。
---

# Markdown 到 Word 高精度转换方案

## 概述

本 Skill 提供基于 **Pandoc + 用户自定义模板** 的 Markdown 到 Word 转换方案，核心特性：

1. **原生公式支持**：LaTeX 公式自动转换为 Word 原生 OMML 公式，完美可编辑
2. **所见即所得样式**：直接在模板 DOCX 中设置字体、编号，无需写代码
3. **智能标题映射**：自动检测 Markdown 标题结构，适配不同写作习惯
4. **中文排版优化**：自动转换中文双引号、清理 CJK/Latin 间距
5. **封面与目录合并**：支持从模板自动合并封面页，标题通过 `{{TITLE}}` 占位符注入
6. **图表题注自动编号**：`图N-M` / `表N-M` 自动转换为 Word SEQ 域，支持自动更新


## 使用场景

当需要进行以下任务时，应激活本 Skill：

- Markdown 转 Word 文档
- 技术报告格式化（含数学公式）
- 学术论文转换
- 需要自定义样式的文档导出

## 依赖环境

| 软件        | 最低版本 | 用途          | 安装命令                               |
| ----------- | -------- | ------------- | -------------------------------------- |
| Pandoc      | 2.x      | 核心转换引擎  | `winget install JohnMacFarlane.Pandoc` |
| Node.js     | 14.x     | 自动化脚本    | `winget install OpenJS.NodeJS.LTS`     |
| Python      | 3.x      | 封面合并      | `winget install Python.Python.3.12`    |
| python-docx | —        | Word 文档操作 | `pip install python-docx`              |

安装验证：
```powershell
pandoc --version
node -v
python --version
```

## 安装与首次配置

### 安装方式

通过 GitHub 一键安装（告诉 AI 即可）：
```powershell
npx @anthropic/skills-cli add https://github.com/wenhaoa/md2word-pandoc
```

安装目标路径：`C:\Users\<用户名>\.gemini\antigravity\skills\md2word-pandoc\`

### 配套组件检查

本 Skill 需要配套的 **写作规则（Rules Part D）** 和 **格式预检工作流（report-check）** 才能发挥完整效果。

**AI 自动检查**：首次激活本 Skill 时，AI 应检查以下条件：
1. 用户的 `GEMINI.md`（全局规则文件）中是否包含 "Part D: Markdown 转 Word 规范"
2. `~/.gemini/antigravity/global_workflows/report-check.md` 是否存在

若缺少以上任一项，AI 应自动读取 [首次安装指南](references/first_time_setup.md)，将其中的 Rules 追加到用户的 `GEMINI.md`，将 Workflow 保存到 `global_workflows/` 目录。

---

## 核心组件

### 1. 转换脚本 (`run_conversion.js`)

**功能**：
- **预处理**：清理 CJK/Latin 间距、转换中文双引号、连接 CJK 跨行软换行
- 自动备份源文件到临时 ASCII 路径（避免中文路径问题）
- 调用 Pandoc 执行转换（禁用 smart 扩展防止引号干扰）
- 合并封面模板与正文（自动替换 `{{TITLE}}` 占位符）
- 生成带北京时间戳的输出文件
- 自动清理临时文件

**关键配置**：
```javascript
const mdFile = '源文件.md';
const referenceDoc = 'md2word模板.docx';  // 样式模板
const filterScript = 'style_filter.lua';   // 标题处理过滤器
```

**时间戳处理**：
```javascript
// 自动修正为北京时间 (+8)
const offset = now.getTimezoneOffset() * 60000;
const localDate = new Date(now.getTime() - offset);
```

---

### 2. 智能标题过滤器 (`style_filter.lua`)

**自动检测逻辑**：

| H1 数量 | 场景描述         | 映射策略                                              |
| ------- | ---------------- | ----------------------------------------------------- |
| **0**   | 文档从 `##` 开始 | `##` → Heading 1，`###` → Heading 2                   |
| **1**   | 有唯一总标题     | `#` 丢弃（由封面 `{{TITLE}}` 展示），`##` → Heading 1 |
| **≥2**  | `#` 作为正式章节 | `#` → Heading 1，`##` → Heading 2                     |

**自动清洗功能**：
- 去除 "第1章"、"1.1" 等手动编号
- 防止与 Word 模板自动编号冲突

**表格样式强制**：
- 所有表格内容应用 `TableContent` 样式
- 便于统一调整表格缩进和间距

---

### 3. 图表题注处理 (`add_captions.py`)

**功能**：Pandoc 转换后自动扫描 `图N-M` / `表N-M` 题注，替换为 Word SEQ 域。

**域代码结构**：
```
图{STYLEREF 1 \s}-{SEQ 图 \* ARABIC \s 1} 标题文本
```
- `STYLEREF 1 \s`：从最近的 Heading 1 获取章节编号（域，自动更新）
- `SEQ 图 \* ARABIC \s 1`：章内自动递增编号，每个 Heading 1 处重置（域，自动更新）

**段落格式**：题注自动设为居中对齐、无首行缩进、段前段后 6pt。

**使用提示**：生成的 docx 打开后按 **Ctrl+A → F9** 更新所有域即可显示正确编号。

---

### 4. 样式模板 (`md2word模板.docx`)

**核心样式定义**：

| 样式名称     | 对应 Markdown | 可自定义项                   |
| ------------ | ------------- | ---------------------------- |
| Heading 1    | `##`（H2）    | 字体、字号、多级列表编号     |
| Heading 2    | `###`（H3）   | 字体、字号、间距             |
| Normal       | 普通段落      | 字体、行距、首行缩进         |
| TableContent | 表格内容      | 字体、缩进（建议设为 0）     |
| Title        | `#`（H1）     | 封面标题样式（仅单 H1 时用） |

**修改样式流程**：
1. 打开 `md2word模板.docx`
2. 在 Word 中修改对应样式（右键 → 修改样式）
3. 保存模板
4. 重新运行转换脚本即可生效

---

## Markdown 文档要求

转 Word 的 Markdown 文件必须遵循以下规范：

### 必须：YAML frontmatter 标题

每个 Markdown 文件**必须**在开头写 YAML frontmatter 的 `title` 字段，该字段用于生成 Word 封面页标题（通过 `{{TITLE}}` 占位符注入）。

```markdown
---
title: 低轨卫星离轨技术研究方案报告
---

## 1. 概述
...
```

**不写 `title` 则 Word 封面标题为空白**，这是最常见的遗漏。

### 标题层级

- 正文从 `##`（H2）开始，**不用** `#`（H1），过滤器自动映射为 Word Heading 1
- 编号带点号：`## 1. 概述`（不是 `## 1 概述`）
- 附录：`## 附录 A 标题`

### 图表题注写法

**图片题注**（紧跟图片下方）：
```markdown
![图2-1 智能温室控制系统架构示意图](images/system_architecture.png)
```

**表格题注**（独立一行，位于表格上方）：
```markdown
表2-1 各层核心组件及功能描述

| 层级 | 核心组件 | 功能描述 |
| ---- | -------- | -------- |
| ...  | ...      | ...      |
```

**注意事项**：
- 编号格式：`图N-M` 或 `表N-M`（N=章节号，M=章内序号），减号分隔
- Markdown 中写的编号仅用于可读性和排序，转 Word 后由 SEQ 域自动更新
- 题注行总长度不超过 60 字符（超长会被识别为正文而非独立题注）
- 正文引用写法不受影响（如「如图2-1所示」不会被转换）

### 正文格式

- 避免不必要的加粗
- 段落优先原则：描述性内容写成连贯段落，减少分点列举
- 数学公式用 `$...$` 和 `$$...$$`，简单数字+单位直写（如 1100km）
- 禁用 GitHub 提示块（`> [!NOTE]`），改为 "注：" 前缀段落
- Mermaid 图表需先转 PNG 再引用

完整写作规范见配套 Rules Part D（在 [first_time_setup.md](references/first_time_setup.md) 中）。

---

## AI 操作指令

收到用户的 md 转 word 请求时，按以下步骤执行：

1. **格式预检**：先执行 `/report-check` 工作流对源文件进行格式检查，若发现问题则先修复再继续
2. **确认源文件路径**：获取用户提供的 `.md` 文件绝对路径
3. **执行转换**：运行以下命令：
   ```powershell
   node "$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\scripts\run_conversion.js" "源文件.md"
   ```
   可选参数：`--no-caption`（跳过题注处理）
4. **报告结果**：告知用户输出文件路径（生成在源文件同目录，格式为 `<源文件名>_时间戳.docx`）
5. **提示更新域**：告知用户打开 Word 后按 Ctrl+A → F9 更新所有域（题注编号等）

> 如果用户已配置 PowerShell Profile，也可以直接使用 `md2word "文件.md"`。

---

## 快速开始

### 方式 1：双击 / 拖拽（推荐，无需命令行）

**首次安装**：双击运行 `scripts\install_shortcuts.bat`，自动在桌面和右键「发送到」菜单创建快捷方式。

**日常使用**：

| 操作                             | 说明                              |
| -------------------------------- | --------------------------------- |
| 双击桌面「md2word」图标          | 弹出文件选择对话框，选择 .md 文件 |
| 拖拽 .md 文件到桌面图标上        | 直接转换                          |
| 右键 .md 文件 → 发送到 → md2word | 在文件资源管理器中直接转换        |

转换完成后**自动打开 Word** 查看效果。

---

### 方式 2：PowerShell 命令

**一次配置，永久使用**

1. **配置 PowerShell Profile**（仅需一次）：
   ```powershell
   notepad $PROFILE
   
   # 添加以下内容到文件末尾：
   function md2word {
       param([Parameter(Mandatory=$true)][string]$mdFile)
       $script = "$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\scripts\run_conversion.js"
       node $script $mdFile
   }
   
   # 保存后重新加载
   . $PROFILE
   ```

2. **日常使用**：
   ```powershell
   # 在任意目录下执行
   md2word "你的文件.md"
   ```



---

### 方式 3：直接调用（无需配置）

在任意目录下执行：

```powershell
node "$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\scripts\run_conversion.js" "你的文件.md"
```

---

### 输出说明

转换完成后，输出文件会生成在**源文件同目录**，文件名格式：
```
<源文件名>_2026-02-05T16-30-00.docx
```

**示例**：
- 输入：`C:\Projects\报告.md`
- 输出：`C:\Projects\报告_2026-02-05T16-30-00.docx`

---

## 样式定制指南

### 修改标题编号

1. 打开 `md2word模板.docx`
2. 选择任意 "标题 1" 段落
3. 右键 → 修改样式 → 格式 → 编号
4. 定义多级列表：
   - 级别 1：章（如 "第1章"）
   - 级别 2：节（如 "1.1"）
   - 级别 3：条（如 "1.1.1"）

### 修改正文格式

修改 `Normal` 样式：
- **字体**：宋体、Times New Roman
- **行距**：1.5 倍行距
- **首行缩进**：2 字符
- **段后间距**：0 pt

### 修复表格缩进

修改 `TableContent` 样式：
1. 右键修改样式 → 格式 → 段落
2. 设置"首行缩进"为 **0**
3. 设置"样式基准"为 **（无样式）**（彻底根除正文缩进干扰）

---

## 常见问题

### 1. 表格框线缺失

**原因**：Word 表格样式未设置默认框线

**解决方法**：
1. 点击任意表格 → 表格设计 (Table Design)
2. 找到 **"普通表格" (Table Normal)** 样式
3. 右键修改 → 格式 → 边框和底纹 → 设置全框线

### 2. 公式字体不一致

**问题**：LaTeX 转换的公式与手动输入的数字字体不同

**解决方法**：
- 全选文档内容，设置字体为 **Cambria Math**
- 或在模板中设置默认数学字体

### 3. 中文路径无法识别

**原因**：Pandoc 在某些系统上不支持中文路径

**解决方法**：
- 脚本已自动将文件复制到临时 ASCII 路径
- 如仍有问题，将整个项目移至纯英文路径

### 4. CJK 与 Latin 间距异常

**问题**：AI 生成的 Markdown 中，中文与数字/英文间有多余空格

**已解决**：`run_conversion.js` 的 `cleanSpaces()` 函数自动清理 CJK/Latin 间距，无需手动处理。

---

## 技术细节与开发备忘

转换流程、Pandoc 命令参数、后续增强方向等技术细节见：[technical_details.md](references/technical_details.md)

## 相关资源

- [Pandoc 官方文档](https://pandoc.org/MANUAL.html)
- [Lua 过滤器指南](https://pandoc.org/lua-filters.html)

## 版本历史

- **V1.3** (2026-03-04)：
  - 新增图表题注自动编号（`add_captions.py`），生成 Word SEQ 域
  - 新增 `--no-caption` 命令行参数
  - 禁用 Pandoc subscript/superscript 扩展（防止 `~` 被解析为下标）
  - 新增 `--resource-path` 修复右键发送图片丢失问题
  - 示例文档增加图片和表格题注演示
- **V1.2** (2026-03-03)：
  - 新增首次安装引导与配套检查（Rules Part D + report-check workflow）
  - 新增 Markdown 文档要求章节（明确 title frontmatter 规范）
  - 转换前自动执行 `/report-check` 格式预检
  - 完善 README.md 展示和安装说明
- **V1.1** (2026-03-02)：
  - 修复封面标题重复渲染问题（清除 `doc.meta.title`，H1 由封面 `{{TITLE}}` 展示）
  - 修复中文双引号方向错误（成对匹配预转换 + 禁用 Pandoc smart 扩展）
  - 新增 YAML frontmatter 保护（防止引号转换破坏 YAML 语法）
  - 新增表格框线与居中自动后处理
- **V1.0** (2026-02-05)：初始版本，包含核心转换流程和智能标题映射

## 维护指南

### 单一真相源

| 信息             | 权威源                                | 同步到                              |
| ---------------- | ------------------------------------- | ----------------------------------- |
| 写作规则 D.1~D.5 | `references/first_time_setup.md`      | → `GEMINI.md` Part D                |
| 预检规则 D.4     | `references/first_time_setup.md` §D.4 | → 全局 `report-check.md` 工作流     |
| 版本历史         | 本文件（SKILL.md）                    | → `README.md`（仅版本号）           |
| 转换流程技术细节 | `run_conversion.js` 源码              | → `references/technical_details.md` |
| 使用方式         | 本文件（SKILL.md）                    | `README.md` 仅简要引用              |

### 新增功能时的更新检查清单

修改转换流程后，按以下清单逐项确认：

1. `scripts/run_conversion.js` — 主流程代码
2. `SKILL.md` — AI 指令（参数说明、操作步骤、版本历史）
3. `references/technical_details.md` — 技术备忘（流程图、命令参数）
4. `references/first_time_setup.md` — 若涉及新的写作规则
5. `GEMINI.md` Part D — 若 first_time_setup 有变更则同步
6. `examples/示例技术报告.md` — 新增功能的演示
7. `README.md` — 仅更新版本号
