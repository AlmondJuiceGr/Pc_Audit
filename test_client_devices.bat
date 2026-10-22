@echo off
cd /d "%~dp0"
py -3 register_client.py --print-diagnostics
if not "%RUN_BACKGROUND%"=="1" pause
