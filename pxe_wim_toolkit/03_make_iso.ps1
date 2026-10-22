param(
    [string]$WinPeRoot = "",
    [string]$IsoPath = ""
)

$ErrorActionPreference = "Stop"

$ToolkitRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$AdkRoot = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit"

function Find-Tool {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [string[]]$FallbackPaths = @()
    )

    $command = Get-Command $Name -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    foreach ($path in $FallbackPaths) {
        if (Test-Path $path) {
            return $path
        }
    }

    return ""
}

if (-not $WinPeRoot) {
    $WinPeRoot = Join-Path $ToolkitRoot "build\WinPE_PCAudit"
}

if (-not $IsoPath) {
    $IsoPath = Join-Path $ToolkitRoot "build\PC_Audit_WinPE.iso"
}

$MediaDir = Join-Path $WinPeRoot "media"
$BootWim = Join-Path $MediaDir "sources\boot.wim"

if (-not (Test-Path $MediaDir)) {
    throw "WinPE media folder not found: $MediaDir. Run 01_build_winpe_audit.ps1 first."
}

if (-not (Test-Path $BootWim)) {
    throw "boot.wim not found: $BootWim. Run 01_build_winpe_audit.ps1 first."
}

if (Test-Path $IsoPath) {
    Write-Host "Removing old ISO: $IsoPath"
    Remove-Item -LiteralPath $IsoPath -Force
}

$oscdimg = Find-Tool "oscdimg.exe" @(
    (Join-Path $AdkRoot "Deployment Tools\amd64\Oscdimg\oscdimg.exe"),
    (Join-Path $AdkRoot "Deployment Tools\x86\Oscdimg\oscdimg.exe"),
    (Join-Path $AdkRoot "Deployment Tools\arm64\Oscdimg\oscdimg.exe")
)

if ($oscdimg) {
    $oscdimgDir = Split-Path -Parent $oscdimg
    if ($env:Path -notlike "*$oscdimgDir*") {
        $env:Path = "$oscdimgDir;$env:Path"
    }
}

$makeWinPeMedia = Find-Tool "MakeWinPEMedia.cmd" @(
    (Join-Path $AdkRoot "Windows Preinstallation Environment\MakeWinPEMedia.cmd")
)
if ($makeWinPeMedia) {
    & $makeWinPeMedia /ISO /f $WinPeRoot $IsoPath
    if ($LASTEXITCODE -ne 0) {
        throw "MakeWinPEMedia failed. Exit code: $LASTEXITCODE"
    }
} else {
    if (-not $oscdimg) {
        throw "Neither MakeWinPEMedia.cmd nor oscdimg.exe was found. Run from Deployment and Imaging Tools Environment."
    }

    $efisys = Join-Path $MediaDir "efi\microsoft\boot\efisys.bin"
    $etfsboot = Join-Path $MediaDir "boot\etfsboot.com"

    if ((Test-Path $efisys) -and (Test-Path $etfsboot)) {
        & $oscdimg -m -o -u2 -udfver102 "-bootdata:2#p0,e,b$etfsboot#pEF,e,b$efisys" $MediaDir $IsoPath
    } else {
        & $oscdimg -m -o -u2 -udfver102 $MediaDir $IsoPath
    }

    if ($LASTEXITCODE -ne 0) {
        throw "oscdimg failed. Exit code: $LASTEXITCODE"
    }
}

Write-Host ""
Write-Host "Done."
Write-Host "ISO: $IsoPath"
