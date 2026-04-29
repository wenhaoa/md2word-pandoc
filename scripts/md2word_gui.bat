@echo off
setlocal
title md2word ????

set "PS_SCRIPT=%~dp0md2word_gui.ps1"

if not exist "%PS_SCRIPT%" (
    echo [Error] md2word_gui.ps1 not found: %PS_SCRIPT%
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*
exit /b %ERRORLEVEL%
