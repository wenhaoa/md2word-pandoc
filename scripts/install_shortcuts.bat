@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title md2word Shortcut Installer

set "BAT_PATH=%~dp0md2word_gui.bat"
set "DESKTOP=%USERPROFILE%\Desktop"
set "SENDTO=%APPDATA%\Microsoft\Windows\SendTo"

echo.
echo ======================================
echo   md2word 快捷方式安装
echo ======================================
echo.

if not exist "!BAT_PATH!" (
    echo [Error] md2word_gui.bat not found
    echo    Path: !BAT_PATH!
    pause
    exit /b 1
)

echo [1/2] 创建桌面快捷方式...
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut('!DESKTOP!\md2word.lnk'); $sc.TargetPath = '!BAT_PATH!'; $sc.WorkingDirectory = '%USERPROFILE%'; $sc.Description = 'Markdown to Word'; $sc.Save()"
if !ERRORLEVEL! equ 0 (
    echo   [OK] 桌面快捷方式已创建
) else (
    echo   [FAIL] 创建失败
)

echo [2/2] 创建「发送到」快捷方式...
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut('!SENDTO!\md2word.lnk'); $sc.TargetPath = '!BAT_PATH!'; $sc.WorkingDirectory = '%USERPROFILE%'; $sc.Description = 'Markdown to Word'; $sc.Save()"
if !ERRORLEVEL! equ 0 (
    echo   [OK]「发送到」快捷方式已创建
) else (
    echo   [FAIL] 创建失败
)

echo.
echo ======================================
echo  安装完成！
echo.
echo  桌面图标：双击选择文件 / 拖拽 .md 文件
echo  发送到：右键 .md - 发送到 - md2word
echo ======================================
echo.
pause
