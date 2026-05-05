@echo off
setlocal

rem Edit SERVER_URL to match your server.
set "SERVER_URL=http://192.168.1.50:8765"
set "CLIENT_NAME=%COMPUTERNAME%"

cd /d "%~dp0"
python register_client.py --server "%SERVER_URL%" --name "%CLIENT_NAME%"

if not "%RUN_BACKGROUND%"=="1" pause
