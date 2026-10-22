@echo off
setlocal

set "CAPTURE_DRIVE=C:\"
set "IMAGE_DIR=D:\Images"
set "IMAGE_FILE=%IMAGE_DIR%\install.wim"
set "IMAGE_NAME=Windows Audit Image"

if not exist "%IMAGE_DIR%" mkdir "%IMAGE_DIR%"

echo Capturing %CAPTURE_DRIVE% to %IMAGE_FILE%
echo If Windows is not on C:, edit CAPTURE_DRIVE in this script.
pause

dism /Capture-Image /ImageFile:"%IMAGE_FILE%" /CaptureDir:"%CAPTURE_DRIVE%" /Name:"%IMAGE_NAME%" /Compress:max /CheckIntegrity

echo.
echo Done.
pause
