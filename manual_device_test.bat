@echo off
setlocal

rem Edit SERVER_URL to match your server.
set "SERVER_URL=http://192.168.1.50:8765"
set "CLIENT_NAME=%COMPUTERNAME%"

cd /d "%~dp0"
if "%RUN_BACKGROUND%"=="1" (
    rem Manual checklist needs visible prompts, so background mode only registers this client safely.
    py -3 register_client.py --server "%SERVER_URL%" --name "%CLIENT_NAME%" --once
    exit /b %ERRORLEVEL%
)

py -3 register_client.py --server "%SERVER_URL%" --name "%CLIENT_NAME%" --manual-tests --once

pause
