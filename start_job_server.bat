@echo off
setlocal

set "JOB_HOST=0.0.0.0"
set "JOB_PORT=8765"

cd /d "%~dp0"
py -3 network_job_server.py --host "%JOB_HOST%" --port "%JOB_PORT%"

if not "%RUN_BACKGROUND%"=="1" pause
