@echo off
setlocal
chcp 65001 >nul 2>&1
title md2word Shortcut Installer

set "NO_PAUSE=0"
if /I "%~1"=="--no-pause" set "NO_PAUSE=1"

set "PS_SCRIPT=%~dp0install_shortcuts.ps1"

echo.
echo ======================================
echo   md2word shortcut install/update
echo ======================================
echo.

if not exist "%PS_SCRIPT%" (
    echo [Error] install_shortcuts.ps1 not found
    echo    Path: %PS_SCRIPT%
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
if errorlevel 1 (
    echo.
    echo [FAIL] shortcut install/update failed.
    if "%NO_PAUSE%"=="0" pause
    exit /b 1
)

echo.
echo Tip: after updating or moving md2word, re-run install_or_update.ps1
echo      to refresh Desktop and SendTo shortcuts.
echo.
if "%NO_PAUSE%"=="0" pause
