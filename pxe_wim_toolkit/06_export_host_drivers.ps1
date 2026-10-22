param(
    [string]$Destination = ""
)

$ErrorActionPreference = "Stop"

$ToolkitRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Destination) {
    $Destination = Join-Path $ToolkitRoot "drivers\host_export"
}

New-Item -ItemType Directory -Force -Path $Destination | Out-Null

Write-Host "Exporting installed third-party drivers to:"
Write-Host $Destination
Write-Host ""

pnputil /export-driver * $Destination
if ($LASTEXITCODE -ne 0) {
    throw "pnputil driver export failed. Exit code: $LASTEXITCODE"
}

Write-Host ""
Write-Host "Done. Rebuild WinPE with:"
Write-Host ".\01_build_winpe_audit.ps1"
