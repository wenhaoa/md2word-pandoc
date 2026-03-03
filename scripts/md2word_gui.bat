@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title md2word 转换工具

REM ============================================================
REM  md2word GUI Launcher
REM  1. Double-click -> File open dialog
REM  2. Drag-drop .md file onto this icon
REM  3. SendTo menu
REM ============================================================

set "SCRIPT_DIR=%~dp0"
set "CONVERSION_SCRIPT=!SCRIPT_DIR!run_conversion.js"

REM 检查 node 是否可用
where node >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo.
    echo [错误] 未找到 Node.js，请先安装
    echo    安装命令：winget install OpenJS.NodeJS.LTS
    echo.
    pause
    exit /b 1
)

REM 检查转换脚本是否存在
if not exist "!CONVERSION_SCRIPT!" (
    echo.
    echo [错误] 转换脚本不存在: !CONVERSION_SCRIPT!
    echo.
    pause
    exit /b 1
)

REM 判断是否有拖拽/SendTo 传入的文件
if "%~1"=="" (
    REM 无参数 → 弹出文件选择对话框
    echo.
    echo 请选择要转换的 Markdown 文件...
    echo.

    for /f "delims=" %%F in ('powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Title = '选择 Markdown 文件'; $d.Filter = 'Markdown 文件 (*.md)|*.md|所有文件 (*.*)|*.*'; $d.FilterIndex = 1; if ($d.ShowDialog() -eq 'OK') { $d.FileName } else { '' }"') do set "MD_FILE=%%F"

    if "!MD_FILE!"=="" (
        echo 已取消。
        timeout /t 2 /nobreak >nul
        exit /b 0
    )
) else (
    REM 拖拽/SendTo 模式：使用传入的文件路径
    set "MD_FILE=%~1"
)

REM 检查文件是否存在
if not exist "!MD_FILE!" (
    echo.
    echo [错误] 文件不存在: !MD_FILE!
    echo.
    pause
    exit /b 1
)

echo.
echo 开始转换: !MD_FILE!
echo =====================================
echo.

REM 调用转换脚本，--open 自动打开生成的 Word 文件
node "!CONVERSION_SCRIPT!" "!MD_FILE!" --open

if !ERRORLEVEL! neq 0 (
    echo.
    echo =====================================
    echo 转换失败，请检查上方错误信息
    echo.
    pause
    exit /b 1
)

echo =====================================
echo 转换完成！Word 文件已自动打开。
echo.

REM 成功时等待 3 秒后自动关闭
timeout /t 3 /nobreak >nul
