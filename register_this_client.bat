@echo off
setlocal

rem Edit SERVER_URL to match your server.
set "SERVER_URL=http://192.168.1.3:8765"
set "CLIENT_NAME=%COMPUTERNAME%"

cd /d "%~dp0"
py -3 register_client.py --server "%SERVER_URL%" --name "%CLIENT_NAME%"

if not "%RUN_BACKGROUND%"=="1" pause
