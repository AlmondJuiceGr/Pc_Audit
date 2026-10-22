param(
    [string]$PythonExe = "",
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$ToolkitRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$PortablePython = Join-Path $ToolkitRoot "portable_python"

if (-not $PythonExe) {
    $PythonExe = (& py -3 -c "import sys; print(sys.executable)").Trim()
}

if (-not (Test-Path $PythonExe)) {
    throw "Python executable not found: $PythonExe"
}

Write-Host "Testing Tkinter..."
& $PythonExe -c "import tkinter; print('tkinter OK')" | Out-Host
if ($LASTEXITCODE -ne 0) {
    throw "This Python does not include Tkinter. Install Python with Tcl/Tk support."
}

$PythonRoot = Split-Path -Parent $PythonExe
Write-Host "Python root: $PythonRoot"
Write-Host "Destination: $PortablePython"

if (Test-Path $PortablePython) {
    if (-not $Force) {
        throw "portable_python already exists. Re-run with -Force to replace it."
    }
    Remove-Item -LiteralPath $PortablePython -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $PortablePython | Out-Null
Get-ChildItem -LiteralPath $PythonRoot -Force | Copy-Item -Destination $PortablePython -Recurse -Force

$CopiedPython = Join-Path $PortablePython "python.exe"
if (-not (Test-Path $CopiedPython)) {
    throw "python.exe was not copied to $PortablePython"
}

Write-Host "Testing copied Python..."
& $CopiedPython -c "import sys, tkinter; print(sys.executable); print('copied tkinter OK')" | Out-Host

Write-Host ""
Write-Host "Done. Rebuild WinPE with:"
Write-Host ".\01_build_winpe_audit.ps1"
