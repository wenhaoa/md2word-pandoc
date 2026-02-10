# Markdown 到 Word 转换函数
# 添加到 PowerShell Profile 以启用 md2word 命令

function md2word {
    <#
    .SYNOPSIS
    将 Markdown 文件转换为 Word 文档
    
    .DESCRIPTION
    使用 Pandoc 将 Markdown 文件转换为带样式的 Word 文档，
    支持 LaTeX 公式、智能标题映射和自定义样式。
    
    .PARAMETER mdFile
    源 Markdown 文件路径（支持相对路径和绝对路径）
    
    .EXAMPLE
    md2word "报告.md"
    
    .EXAMPLE
    md2word "C:\Projects\文档\技术报告.md"
    #>
    
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$mdFile
    )
    
    # Skill 中的转换脚本路径
    $script = "$env:USERPROFILE\.gemini\antigravity\skills\md2word-pandoc\scripts\run_conversion.js"
    
    # 检查脚本是否存在
    if (-not (Test-Path $script)) {
        Write-Error "❌ 转换脚本不存在: $script"
        Write-Error "   请确认 md2word-pandoc Skill 已正确安装"
        return
    }
    
    # 调用转换脚本
    node $script $mdFile
}

Write-Host "✅ md2word 函数已加载" -ForegroundColor Green
Write-Host "   用法: md2word '文件.md'" -ForegroundColor Gray
