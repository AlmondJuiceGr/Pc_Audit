param()

Write-Host "Run these in Administrator PowerShell on the PXE/WDS server if needed:"
Write-Host ""
Write-Host 'New-NetFirewallRule -DisplayName "PC Audit Job Server 8765" -Direction Inbound -Action Allow -Protocol TCP -LocalPort 8765 -Profile Private,Domain'
Write-Host 'New-NetFirewallRule -DisplayName "PXE DHCP Proxy" -Direction Inbound -Action Allow -Protocol UDP -LocalPort 67,68,69,4011 -Profile Private,Domain'
Write-Host ""
Write-Host "Typical DHCP options if router/DHCP is separate from WDS:"
Write-Host "Option 66: PXE/WDS server IP"
Write-Host "Option 67 BIOS: boot\x64\wdsnbp.com"
Write-Host "Option 67 UEFI: boot\x64\wdsmgfw.efi"
