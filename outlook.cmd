@echo off
setlocal
set "OUTLOOK_NODE=node"
where node >nul 2>&1
if errorlevel 1 set "OUTLOOK_NODE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
"%OUTLOOK_NODE%" "%~dp0bin\outlook.mjs" %*
