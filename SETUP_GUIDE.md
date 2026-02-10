# md2word 配置指南

## 🚀 一次配置，永久使用

### 步骤 1：打开 PowerShell Profile

```powershell
notepad $PROFILE
```

如果提示文件不存在，执行：
```powershell
New-Item -Path $PROFILE -ItemType File -Force
notepad $PROFILE
```

### 步骤 2：添加 md2word 函数

将以下内容复制到文件末尾：

```powershell
# ========== Markdown 到 Word 转换函数 ==========
function md2word {
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$mdFile
    )
    
    $script = "$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\scripts\run_conversion.js"
    
    if (-not (Test-Path $script)) {
        Write-Error "❌ 转换脚本不存在: $script"
        Write-Error "   请确认 md2word-pandoc Skill 已正确安装"
        return
    }
    
    node $script $mdFile
}
```

### 步骤 3：重新加载 Profile

```powershell
. $PROFILE
```

或者重启 PowerShell。

---

## 📖 使用方法

配置完成后，在任意目录下使用：

```powershell
# 相对路径
md2word "报告.md"

# 绝对路径
md2word "C:\Projects\文档\技术报告.md"
```

输出文件自动生成在源文件目录，格式：
```
<源文件名>_Final_<时间戳>.docx
```

---

## ✅ 验证安装

运行以下命令测试：

```powershell
Get-Command md2word
```

如果显示函数定义，说明配置成功！

---

## 🔧 故障排查

**问题 1：找不到 md2word 命令**

解决方法：
1. 确认 Profile 已保存
2. 执行 `. $PROFILE` 重新加载
3. 重启 PowerShell

**问题 2：提示脚本不存在**

解决方法：
确认 Skill 目录存在：
```powershell
Test-Path "$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\scripts\run_conversion.js"
```

**问题 3：Pandoc 错误**

解决方法：
确认 Pandoc 已安装：
```powershell
pandoc --version
```

---

## 💡 高级用法

### 自定义模板

如果需要项目特定的样式模板：

1. 复制默认模板到项目目录：
   ```powershell
   Copy-Item "$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\templates\md2word模板.docx" .
   ```

2. 修改 `md2word模板.docx` 中的样式

3. 再次运行 `md2word "文件.md"`

脚本会优先使用当前目录的模板！

---

## 📚 更多信息

- 完整文档：`$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\SKILL.md`
- Workflow：输入 `/md2word` 查看
