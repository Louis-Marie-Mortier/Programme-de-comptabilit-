@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0LancerFaisonsLesComptes.ps1"
set "CODE=%ERRORLEVEL%"

echo.
if not "%CODE%"=="0" (
    echo Le lancement a signale une erreur.
) else (
    echo Appuyez sur une touche pour fermer cette fenetre.
)
pause >nul

exit /b %CODE%
