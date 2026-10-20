@echo off
setlocal
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
set "OUTLOOK_INSTALL_EXIT=%ERRORLEVEL%"
echo.
if not "%OUTLOOK_INSTALL_EXIT%"=="0" echo Installation failed. Read the error above.
pause
exit /b %OUTLOOK_INSTALL_EXIT%
