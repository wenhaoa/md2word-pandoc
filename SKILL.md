---
name: md2word-pandoc
description: 使用 Pandoc 将 Markdown 高精度转换为 Word 文档，支持原生公式、自定义样式和智能标题映射
---

# Markdown 到 Word 高精度转换方案

## 概述

本 Skill 提供基于 **Pandoc + 用户自定义模板** 的 Markdown 到 Word 转换方案，核心特性：

1. **原生公式支持**：LaTeX 公式自动转换为 Word 原生 OMML 公式，完美可编辑
2. **所见即所得样式**：直接在模板 DOCX 中设置字体、编号，无需写代码
3. **智能标题映射**：自动检测 Markdown 标题结构，适配不同写作习惯
4. **中文排版优化**：自动转换中文双引号、清理 CJK/Latin 间距
5. **封面与目录合并**：支持从模板自动合并封面页，标题通过 `{{TITLE}}` 占位符注入
6. **自动化流程**：处理中文文件名、时间戳、备份等繁琐操作

## 使用场景

当需要进行以下任务时，应激活本 Skill：

- Markdown 转 Word 文档
- 技术报告格式化（含数学公式）
- 学术论文转换
- 需要自定义样式的文档导出

## 依赖环境

- **Pandoc** >= 2.x：核心转换引擎
- **Node.js** >= 14.x：自动化脚本运行环境
- **操作系统**：支持 Windows/Linux/macOS

安装验证：
```powershell
pandoc --version
node -v
```

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

### 3. 样式模板 (`md2word模板.docx`)

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

## 快速开始

### 方式 1：PowerShell 命令（推荐）

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

> 详细配置说明见：[SETUP_GUIDE.md](./SETUP_GUIDE.md)

---

### 方式 2：直接调用（无需配置）

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

## 技术原理

### Pandoc 转换流程

```mermaid
graph LR
    A[源 MD 文件] --> B["预处理（空格清理+引号转换）"]
    B --> C[复制到临时 ASCII 路径]
    C --> D["Pandoc 解析（smart 已禁用）"]
    D --> E[Lua 过滤器处理]
    E --> F[应用模板样式]
    F --> G[合并封面与目录]
    G --> H[重命名为带时间戳文件名]
```

### 核心 Pandoc 命令

```powershell
pandoc 源文件.md \
    -o 输出.docx \
    --from markdown-smart \
    --reference-doc=md2word模板.docx \
    --lua-filter=style_filter.lua \
    --standalone
```

**参数说明**：
- `--from markdown-smart`：禁用 smart typography，防止干扰预处理后的中文引号
- `--reference-doc`：指定样式模板
- `--lua-filter`：应用自定义处理逻辑
- `--standalone`：生成完整文档（含元数据）

---

## 后续增强方向

### 1. 表格后处理脚本

**目标**：自动修复表格框线和缩进

**实现思路**：
- 调用 `docx` Skill
- 解析 XML 批量修改表格样式属性

### ~~2. 空格清理预处理~~ ✅ 已实现

已在 V1.1 中实现，见 `run_conversion.js` 的 `cleanSpaces()` 函数。

### 3. LaTeX 滥用检测

**目标**：检测并清理不必要的 LaTeX 语法（如 `$100$` 应写为 `100`）

**实现思路**：
- 正则匹配纯数字/简单单位的 LaTeX 块
- 自动降级为普通文本

---

## 相关资源

- **GitHub 仓库**：https://github.com/wenhaoa/md2word-pandoc
- **Pandoc 官方文档**：https://pandoc.org/MANUAL.html
- **Lua 过滤器指南**：https://pandoc.org/lua-filters.html
- **Word OOXML 规范**：https://docs.microsoft.com/openxml

---

## 版本历史

- **V1.1** (2026-03-02)：
  - 修复封面标题重复渲染问题（清除 `doc.meta.title`，H1 由封面 `{{TITLE}}` 展示）
  - 修复中文双引号方向错误（成对匹配预转换 + 禁用 Pandoc smart 扩展）
  - 新增 YAML frontmatter 保护（防止引号转换破坏 YAML 语法）
  - 新增表格框线与居中自动后处理
- **V1.0** (2026-02-05)：初始版本，包含核心转换流程和智能标题映射
