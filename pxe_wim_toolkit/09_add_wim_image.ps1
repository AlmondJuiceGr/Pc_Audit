param(
    [Parameter(Mandatory=$true)]
    [string]$SourceWim,
    [string]$Name = "install.wim"
)

$ErrorActionPreference = "Stop"

$ToolkitRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ImagesRoot = Join-Path $ToolkitRoot "images"

if (-not (Test-Path $SourceWim)) {
    throw "WIM file not found: $SourceWim"
}

if ([IO.Path]::GetExtension($Name).ToLowerInvariant() -ne ".wim") {
    $Name = "$Name.wim"
}

New-Item -ItemType Directory -Force -Path $ImagesRoot | Out-Null
$Destination = Join-Path $ImagesRoot $Name

Copy-Item -LiteralPath $SourceWim -Destination $Destination -Force

Write-Host "Added WIM image:"
Write-Host $Destination
Write-Host ""
Write-Host "Rebuild WinPE/ISO:"
Write-Host ".\01_build_winpe_audit.ps1 -PackageSet Gui"
Write-Host ".\03_make_iso.ps1"
