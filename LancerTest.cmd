@echo off
setlocal
cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0LancerTest.ps1"
set "CODE=%ERRORLEVEL%"

if not "%CODE%"=="0" (
    echo.
    echo Le test a signale une erreur. Consulte les messages ci-dessus.
    pause
)

exit /b %CODE%
