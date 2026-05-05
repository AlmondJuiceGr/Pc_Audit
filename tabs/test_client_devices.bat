@echo off
cd /d "%~dp0"
python register_client.py --print-diagnostics
if not "%RUN_BACKGROUND%"=="1" pause
