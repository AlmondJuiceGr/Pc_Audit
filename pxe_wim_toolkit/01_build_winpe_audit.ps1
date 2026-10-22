param(
    [string]$Architecture = "amd64",
    [string]$ServerUrl = "http://192.168.1.3:8765",
    [string]$ClientName = "%COMPUTERNAME%",
    [string]$DriverRoot = "",
    [ValidateSet("Gui", "Network", "Full")]
    [string]$PackageSet = "Gui",
    [switch]$ForceUnsignedDrivers
)

$ErrorActionPreference = "Stop"

$ToolkitRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ToolkitRoot
$BuildRoot = Join-Path $ToolkitRoot "build"
$WinPeRoot = Join-Path $BuildRoot "WinPE_PCAudit"
$MountDir = Join-Path $WinPeRoot "mount"
$MediaDir = Join-Path $WinPeRoot "media"
$BootWim = Join-Path $MediaDir "sources\boot.wim"
$AuditDir = Join-Path $MountDir "PCAudit"
if (-not $DriverRoot) {
    $DriverRoot = Join-Path $ToolkitRoot "drivers"
}

function Invoke-Checked {
    param(
        [Parameter(Mandatory=$true)]
        [scriptblock]$Command,
        [Parameter(Mandatory=$true)]
        [string]$Message
    )

    & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "$Message Exit code: $LASTEXITCODE"
    }
}

function Add-WinPePackageIfPresent {
    param(
        [Parameter(Mandatory=$true)]
        [string]$PackageName
    )

    $ocRoot = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\$Architecture\WinPE_OCs"
    $cab = Join-Path $ocRoot "$PackageName.cab"
    if (-not (Test-Path $cab)) {
        Write-Host "Skipping missing WinPE package: $PackageName"
        return
    }

    Write-Host "Adding WinPE package: $PackageName"
    Invoke-Checked -Message "Failed to add WinPE package $PackageName." -Command {
        dism /Image:$MountDir /Add-Package /PackagePath:$cab | Out-Host
    }

    $langCab = Join-Path $ocRoot "en-us\$($PackageName)_en-us.cab"
    if (Test-Path $langCab) {
        Write-Host "Adding WinPE language package: $PackageName en-us"
        Invoke-Checked -Message "Failed to add WinPE language package $PackageName." -Command {
            dism /Image:$MountDir /Add-Package /PackagePath:$langCab | Out-Host
        }
    }
}

$copype = Get-Command "copype.cmd" -ErrorAction SilentlyContinue
if (-not $copype) {
    throw "copype.cmd not found. Run this from 'Deployment and Imaging Tools Environment' after installing Windows ADK + WinPE add-on."
}

if (Test-Path $WinPeRoot) {
    Write-Host "Removing old build folder: $WinPeRoot"
    Remove-Item -LiteralPath $WinPeRoot -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $BuildRoot | Out-Null
Invoke-Checked -Message "copype failed." -Command { & $copype.Source $Architecture $WinPeRoot }

Write-Host "Mounting $BootWim"
Invoke-Checked -Message "DISM failed to mount the WinPE image. Run this script from an elevated Deployment and Imaging Tools Environment." -Command {
    dism /Mount-Image /ImageFile:$BootWim /Index:1 /MountDir:$MountDir | Out-Host
}

Write-Host "WinPE package set: $PackageSet"
$packages = @("WinPE-WMI", "WinPE-Scripting")
if ($PackageSet -in @("Network", "Full")) {
    $packages += @("WinPE-Dot3Svc", "WinPE-RNDIS")
}
if ($PackageSet -eq "Full") {
    $packages += @(
        "WinPE-NetFX",
        "WinPE-PowerShell",
        "WinPE-StorageWMI",
        "WinPE-DismCmdlets",
        "WinPE-WDS-Tools",
        "WinPE-WiFi-Package"
    )
}

foreach ($package in $packages) {
    Add-WinPePackageIfPresent $package
}

if (Test-Path $DriverRoot) {
    Write-Host "Injecting drivers from: $DriverRoot"
    if ($ForceUnsignedDrivers) {
        Write-Host "Unsigned drivers are allowed for this build."
        Invoke-Checked -Message "Failed to inject drivers from $DriverRoot." -Command {
            dism /Image:$MountDir /Add-Driver /Driver:$DriverRoot /Recurse /ForceUnsigned | Out-Host
        }
    } else {
        Invoke-Checked -Message "Failed to inject drivers from $DriverRoot. If you trust these drivers, re-run with -ForceUnsignedDrivers." -Command {
            dism /Image:$MountDir /Add-Driver /Driver:$DriverRoot /Recurse | Out-Host
        }
    }
} else {
    Write-Host "Driver folder not found; skipping driver injection: $DriverRoot"
}

New-Item -ItemType Directory -Force -Path $AuditDir | Out-Null

$files = @(
    "pc_audit_app.py",
    "pc_audit_app.pyw",
    "app.db",
    "network_jobs.db",
    "register_client.py",
    "pc_audit_report.py",
    "pc_audit_settings.json",
    "requirements.txt",
    "start_pc_audit_app.bat",
    "start_pc_audit.bat",
    "windows_device_tests.bat",
    "register_this_client.bat"
)

foreach ($file in $files) {
    $source = Join-Path $RepoRoot $file
    if (Test-Path $source) {
        Copy-Item -LiteralPath $source -Destination $AuditDir -Force
    }
}

$portablePython = Join-Path $ToolkitRoot "portable_python"
if (Test-Path $portablePython) {
    Copy-Item -LiteralPath $portablePython -Destination (Join-Path $AuditDir "python") -Recurse -Force
}

Copy-Item -LiteralPath (Join-Path $ToolkitRoot "capture_wim.cmd") -Destination $AuditDir -Force
Copy-Item -LiteralPath (Join-Path $ToolkitRoot "apply_wim.cmd") -Destination $AuditDir -Force
Copy-Item -LiteralPath (Join-Path $ToolkitRoot "winpe_launcher.py") -Destination $AuditDir -Force

$winPeLauncher = @"
@echo off
cd /d X:\PCAudit
if not exist X:\PCAudit\python\python.exe (
  echo Portable Python was not found at X:\PCAudit\python\python.exe
  echo Copy portable Python into pxe_wim_toolkit\portable_python and rebuild WinPE.
  pause
  exit /b 1
)

set PYTHONHOME=X:\PCAudit\python
set PYTHONPATH=X:\PCAudit\python\Lib;X:\PCAudit\python\DLLs
set TCL_LIBRARY=X:\PCAudit\python\tcl\tcl8.6
set TK_LIBRARY=X:\PCAudit\python\tcl\tk8.6
set PATH=X:\PCAudit\python;X:\PCAudit\python\DLLs;%PATH%
set LOG=X:\PCAudit\winpe_app_error.log
set PC_AUDIT_SERVER_URL=$ServerUrl
set PC_AUDIT_CLIENT_NAME=$ClientName-WinPE

if exist X:\PCAudit\winpe_launcher.py (
  X:\PCAudit\python\python.exe X:\PCAudit\winpe_launcher.py 1>%LOG% 2>&1
  echo.
  echo GUI launcher closed or failed. Terminal fallback menu is starting.
  if exist %LOG% type %LOG%
)

echo.
echo PC Audit WinPE Menu
echo ====================
echo 1. Run PC Audit Tool
echo 2. Run client registration audit once
echo 3. Run full PC Audit app
echo 4. Capture Windows to WIM
echo 5. Apply WIM to disk
echo 6. Diagnose Python/Tkinter
echo 7. Command prompt
echo.
choice /c 1234567 /n /m "Select option: "

if errorlevel 7 cmd
if errorlevel 6 (
  echo.
  echo Running diagnosis...
  X:\PCAudit\python\python.exe -c "import sys, os, tkinter; print(sys.executable); print('tkinter OK'); print(os.getcwd())"
  echo.
  dir X:\PCAudit
  echo.
  pause
  X:\PCAudit\start_pc_audit_winpe.cmd
)
if errorlevel 5 call X:\PCAudit\apply_wim.cmd
if errorlevel 4 call X:\PCAudit\capture_wim.cmd
if errorlevel 3 X:\PCAudit\python\python.exe X:\PCAudit\pc_audit_app.pyw 1>%LOG% 2>&1
if errorlevel 2 X:\PCAudit\python\python.exe X:\PCAudit\register_client.py --server "$ServerUrl" --name "$ClientName-WinPE" --once 1>%LOG% 2>&1
if errorlevel 1 X:\PCAudit\python\python.exe X:\PCAudit\pc_audit_report.py 1>%LOG% 2>&1

echo.
if exist %LOG% (
  echo If the program did not open, check this log:
  echo   %LOG%
  type %LOG%
)
pause
X:\PCAudit\start_pc_audit_winpe.cmd
"@

Set-Content -LiteralPath (Join-Path $AuditDir "start_pc_audit_winpe.cmd") -Value $winPeLauncher -Encoding ASCII

$startnet = @"
wpeinit
wpeutil InitializeNetwork
echo.
echo PC Audit WinPE Environment
echo Server: $ServerUrl
ipconfig
echo.
if exist X:\PCAudit\python\python.exe (
  call X:\PCAudit\start_pc_audit_winpe.cmd
) else (
  echo Portable Python was not found at X:\PCAudit\python\python.exe
  echo Copy portable Python into pxe_wim_toolkit\portable_python and rebuild WinPE.
)
echo.
echo Available commands:
echo   X:\PCAudit\start_pc_audit_winpe.cmd
echo   X:\PCAudit\capture_wim.cmd
echo   X:\PCAudit\apply_wim.cmd
echo.
cmd
"@

$startnetPath = Join-Path $MountDir "Windows\System32\startnet.cmd"
Set-Content -LiteralPath $startnetPath -Value $startnet -Encoding ASCII

Write-Host "Unmounting and committing WinPE image"
Invoke-Checked -Message "DISM failed to unmount and commit the WinPE image." -Command {
    dism /Unmount-Image /MountDir:$MountDir /Commit | Out-Host
}

$ImagesRoot = Join-Path $ToolkitRoot "images"
if (Test-Path $ImagesRoot) {
    $MediaImages = Join-Path $MediaDir "Images"
    Write-Host "Copying WIM images into ISO media: $MediaImages"
    New-Item -ItemType Directory -Force -Path $MediaImages | Out-Null
    Copy-Item -LiteralPath (Join-Path $ImagesRoot "*") -Destination $MediaImages -Recurse -Force
}

Write-Host ""
Write-Host "Done."
Write-Host "Boot WIM: $BootWim"
