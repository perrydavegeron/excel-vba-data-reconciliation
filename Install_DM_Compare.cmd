@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build_and_Install_DM_Compare.ps1"
echo.
pause
