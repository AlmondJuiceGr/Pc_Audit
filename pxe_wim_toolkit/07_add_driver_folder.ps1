param(
    [Parameter(Mandatory=$true)]
    [string]$Source,
    [string]$Name = ""
)

$ErrorActionPreference = "Stop"

$ToolkitRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$DriversRoot = Join-Path $ToolkitRoot "drivers"

if (-not (Test-Path $Source)) {
    throw "Driver source folder not found: $Source"
}

$infFiles = Get-ChildItem -LiteralPath $Source -Recurse -Filter *.inf -File
if (-not $infFiles) {
    throw "No .inf driver files were found in: $Source"
}

if (-not $Name) {
    $Name = Split-Path -Leaf (Resolve-Path $Source)
}

$safeName = ($Name -replace '[^\w\-. ]', '_').Trim()
if (-not $safeName) {
    $safeName = "custom_drivers"
}

$Destination = Join-Path $DriversRoot $safeName

if (Test-Path $Destination) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $Destination = Join-Path $DriversRoot "$safeName`_$timestamp"
}

New-Item -ItemType Directory -Force -Path $Destination | Out-Null
Copy-Item -LiteralPath $Source -Destination $Destination -Recurse -Force

Write-Host "Added driver folder:"
Write-Host $Destination
Write-Host ""
Write-Host "INF files found: $($infFiles.Count)"
Write-Host ""
Write-Host "Rebuild WinPE to inject these drivers:"
Write-Host ".\01_build_winpe_audit.ps1"
