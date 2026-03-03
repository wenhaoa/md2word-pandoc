# 技术原理与开发备忘

## Pandoc 转换流程

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

## 核心 Pandoc 命令

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
