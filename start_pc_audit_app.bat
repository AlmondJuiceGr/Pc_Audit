@echo off
setlocal

cd /d "%~dp0"
py -3 pc_audit_app.py

if not "%RUN_BACKGROUND%"=="1" pause
