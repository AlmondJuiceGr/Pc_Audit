@echo off
setlocal

rem Required Windows tests:
rem - USB device detection/count
rem - Optical read test
rem - USB/SD removable read test
rem - Sound left/right channel test

rem Edit SERVER_URL or set JOB_SERVER_URL before running.
set "SERVER_URL=%JOB_SERVER_URL%"
if "%SERVER_URL%"=="" set "SERVER_URL=http://192.168.1.50:8765"

set "CLIENT_NAME=%CLIENT_NAME%"
if "%CLIENT_NAME%"=="" set "CLIENT_NAME=%COMPUTERNAME%"

cd /d "%~dp0"
set "EXTRA_ARGS="
if "%RUN_BACKGROUND%"=="1" set "EXTRA_ARGS=--non-interactive-tests"
py -3 register_client.py --server "%SERVER_URL%" --name "%CLIENT_NAME%" --windows-device-tests --no-diagnostics %EXTRA_ARGS% --once

if not "%RUN_BACKGROUND%"=="1" pause
