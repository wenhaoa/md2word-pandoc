---
description: 使用 Pandoc 将 Markdown 转换为样式完整的 Word 文档
---

# Markdown 到 Word 转换流程

## ⚡ 快速使用（推荐）

### 1. 首次配置（仅需一次）

将以下内容添加到 PowerShell Profile：

```powershell
# 编辑 Profile
notepad $PROFILE

# 添加以下内容到文件末尾：
function md2word {
    param([Parameter(Mandatory=$true)][string]$mdFile)
    $script = "$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\scripts\run_conversion.js"
    node $script $mdFile
}
```

保存后重启 PowerShell 或执行 `. $PROFILE` 重新加载。

### 2. 日常使用

在任意目录下，直接运行：

// turbo
```powershell
md2word "你的文件.md"
```

**示例**：
```powershell
md2word "高精度轨道递推报告.md"
md2word "C:\Projects\技术文档\研究报告.md"
```

输出文件会自动生成在源文件同目录，文件名格式：
```
源文件名_Final_2026-02-05T16-30-00.docx
```

---

## 🔧 手动调用（无需配置）

如果不想配置 PowerShell 函数，可以直接调用脚本：

```powershell
node "$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\scripts\run_conversion.js" "你的文件.md"
```

---

## 📋 后续处理

### 表格格式修复

转换后表格可能缺少框线或有缩进问题，需在 Word 中手动修复：

1. **修复框线**：
   - 点击任意表格 → 表格设计 (Table Design)
   - 找到 "普通表格" (Table Normal) 样式
   - 右键修改 → 格式 → 边框和底纹 → 设置全框线

2. **修复缩进**：
   - 修改 `TableContent` 样式
   - 设置"首行缩进"为 **0**
   - 设置"样式基准"为 **（无样式）**

---

## ❓ 常见问题

**Q: 标题编号没有出现？**  
A: 这是模板问题。如需自定义：
1. 复制模板到项目目录：`Copy-Item "$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\templates\md2word模板.docx" .`
2. 修改模板中的 "Heading 1" 样式，关联多级列表
3. 再次运行转换（脚本会优先使用当前目录的模板）

**Q: 公式显示异常？**  
A: 全选文档内容，设置字体为 **Cambria Math**。

**Q: 找不到 md2word 命令？**  
A: 请确认 PowerShell Profile 已添加函数，并重启 PowerShell。

---

## 🎯 核心优势

- ✅ **零文件复制**：无需复制任何配置文件
- ✅ **一键转换**：`md2word "文件.md"` 即可
- ✅ **原生公式**：LaTeX 自动转为 Word 公式
- ✅ **智能标题**：自动检测 H1 数量，适配不同结构
- ✅ **统一管理**：升级 Skill 即升级转换工具

---

## 📚 相关说明

详细技术说明和样式定制方法，请参考：  
`$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\SKILL.md`
