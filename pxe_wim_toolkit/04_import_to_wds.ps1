param(
    [Parameter(Mandatory=$true)]
    [string]$BootWim,

    [Parameter(Mandatory=$false)]
    [string]$InstallWim = "",

    [string]$ImageGroup = "PC Audit Images"
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command wdsutil.exe -ErrorAction SilentlyContinue)) {
    throw "wdsutil.exe not found. Install/configure Windows Deployment Services first."
}

if (-not (Test-Path $BootWim)) {
    throw "Boot WIM not found: $BootWim"
}

Write-Host "Adding boot image: $BootWim"
wdsutil /Add-Image /ImageFile:"$BootWim" /ImageType:Boot | Out-Host

if ($InstallWim) {
    if (-not (Test-Path $InstallWim)) {
        throw "Install WIM not found: $InstallWim"
    }

    Write-Host "Ensuring image group: $ImageGroup"
    wdsutil /Add-ImageGroup /ImageGroup:"$ImageGroup" | Out-Host

    Write-Host "Adding install image: $InstallWim"
    wdsutil /Add-Image /ImageFile:"$InstallWim" /ImageType:Install /ImageGroup:"$ImageGroup" | Out-Host
}

Write-Host "Done."
