param()

$ErrorActionPreference = "Continue"

function Test-Command($Name) {
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if ($cmd) {
        [pscustomobject]@{ Tool = $Name; Found = $true; Path = $cmd.Source }
    } else {
        [pscustomobject]@{ Tool = $Name; Found = $false; Path = "" }
    }
}

$tools = @(
    Test-Command "dism.exe"
    Test-Command "copype.cmd"
    Test-Command "MakeWinPEMedia.cmd"
    Test-Command "oscdimg.exe"
    Test-Command "wdsutil.exe"
)

$tools | Format-Table -AutoSize

$adkRoots = @(
    "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit",
    "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit"
)

Write-Host ""
Write-Host "ADK folders:"
foreach ($root in $adkRoots) {
    if (Test-Path $root) {
        Write-Host "FOUND $root"
    } else {
        Write-Host "MISSING $root"
    }
}

Write-Host ""
Write-Host "Current user is administrator:"
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

Write-Host ""
Write-Host "If copype.cmd is missing, install Windows ADK and Windows PE add-on."
