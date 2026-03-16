@echo off
setlocal
set SCRIPT_DIR=%~dp0
start "" "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File "%SCRIPT_DIR%Uchatgpt.ps1"
endlocal
