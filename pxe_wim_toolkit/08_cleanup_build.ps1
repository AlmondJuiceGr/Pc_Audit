param(
    [switch]$RemoveIso
)

$ErrorActionPreference = "Continue"

$ToolkitRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$BuildRoot = Join-Path $ToolkitRoot "build"
$WinPeRoot = Join-Path $BuildRoot "WinPE_PCAudit"
$MountDir = Join-Path $WinPeRoot "mount"
$IsoPath = Join-Path $BuildRoot "PC_Audit_WinPE.iso"

Write-Host "Cleaning WinPE build state..."

if (Test-Path $MountDir) {
    Write-Host "Trying DISM unmount/discard:"
    Write-Host $MountDir
    dism /Unmount-Image /MountDir:$MountDir /Discard | Out-Host
}

Write-Host "Cleaning stale DISM mounts..."
dism /Cleanup-Mountpoints | Out-Host

if ($RemoveIso -and (Test-Path $IsoPath)) {
    Write-Host "Removing old ISO:"
    Write-Host $IsoPath
    Remove-Item -LiteralPath $IsoPath -Force
}

if (Test-Path $WinPeRoot) {
    Write-Host "Removing old WinPE build folder:"
    Write-Host $WinPeRoot
    try {
        Remove-Item -LiteralPath $WinPeRoot -Recurse -Force
    } catch {
        Write-Host ""
        Write-Host "Could not remove the build folder."
        Write-Host "Close File Explorer windows, unmount any mounted ISO, close apps using boot.wim, then run this script again."
        throw
    }
}

Write-Host ""
Write-Host "Cleanup done."
