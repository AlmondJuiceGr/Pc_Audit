@echo off
setlocal

set "IMAGE_FILE="
set "APPLY_DRIVE=C:\"

if exist X:\PCAudit\Images\install.wim set "IMAGE_FILE=X:\PCAudit\Images\install.wim"
if not defined IMAGE_FILE (
  for %%D in (C D E F G H I J K L M N O P Q R S T U V W Y Z) do (
    if exist %%D:\Images\install.wim set "IMAGE_FILE=%%D:\Images\install.wim"
  )
)
if not defined IMAGE_FILE (
  for %%D in (C D E F G H I J K L M N O P Q R S T U V W Y Z) do (
    for %%F in (%%D:\Images\*.wim) do (
      if not defined IMAGE_FILE set "IMAGE_FILE=%%F"
    )
  )
)
if not defined IMAGE_FILE (
  echo No WIM image found.
  echo Put your image at \Images\install.wim on the boot ISO/USB or on another drive.
  pause
  exit /b 1
)

echo WARNING: This applies a Windows image to %APPLY_DRIVE%.
echo Make sure the disk is partitioned correctly before continuing.
echo.
echo Image: %IMAGE_FILE%
echo Target: %APPLY_DRIVE%
echo.
set /p CONFIRM=Type APPLY to continue: 
if /i not "%CONFIRM%"=="APPLY" exit /b 1

dism /Apply-Image /ImageFile:"%IMAGE_FILE%" /Index:1 /ApplyDir:"%APPLY_DRIVE%"
bcdboot %APPLY_DRIVE%\Windows

echo.
echo Done.
pause
