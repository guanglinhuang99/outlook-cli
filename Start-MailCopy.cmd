@echo off
setlocal
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0powershell\clipboard-control.ps1" -Action start -Quiet
if errorlevel 1 (
    echo Could not start Outlook Mail Copy. See the message above.
    pause
    exit /b 1
)
exit /b 0
