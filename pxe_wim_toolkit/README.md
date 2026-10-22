# PXE / WIM Toolkit for PC Audit

This folder contains helper scripts for a Windows deployment flow:

1. Build a WinPE boot image.
2. Copy the PC Audit files into WinPE.
3. Capture a prepared Windows machine to a `.wim`.
4. Apply a `.wim` from WinPE.
5. Import boot/install images into Windows Deployment Services.

## Requirements

- Windows ADK
- Windows PE add-on for the ADK
- DISM
- Administrator PowerShell
- WDS/MDT or another PXE server

Run this first:

```powershell
.\00_check_requirements.ps1
```

## Build WinPE Audit Image

Open **Deployment and Imaging Tools Environment as Administrator**. From the cloned repository, enter the toolkit folder, then run:

```powershell
cd "<path-to-repository>\pxe_wim_toolkit"
.\01_build_winpe_audit.ps1
```

Default package mode is `Gui`. It keeps the Python/Tkinter GUI launcher and basic
hardware inventory support while avoiding heavy optional packages.

Package modes:

```powershell
.\01_build_winpe_audit.ps1 -PackageSet Gui
.\01_build_winpe_audit.ps1 -PackageSet Network
.\01_build_winpe_audit.ps1 -PackageSet Full
```

- `Gui`: GUI launcher, Python/Tkinter app, basic WMI/Scripting.
- `Network`: `Gui` plus wired network optional components.
- `Full`: more Microsoft optional packages, but slower and more likely to hang on some ADK builds.

Output:

```text
build\WinPE_PCAudit\media\sources\boot.wim
```

Import that `boot.wim` into WDS as a boot image.

## Python in WinPE

WinPE does not include Python by default.

If you want the audit script or the GUI tools to run automatically inside WinPE, place a portable Python folder here:

```text
pxe_wim_toolkit\portable_python\
```

You can create that folder from your installed Python with:

```powershell
.\02_prepare_portable_python.ps1
```

If the folder already exists and you want to replace it:

```powershell
.\02_prepare_portable_python.ps1 -Force
```

The build script copies it into WinPE as:

```text
X:\PCAudit\python\
```

The generated WinPE startup script tries:

```text
X:\PCAudit\python\python.exe X:\PCAudit\register_client.py ...
```

If Python is not included, WinPE opens a command prompt and shows instructions.

For the full GUI tools, the portable Python build must include Tkinter/Tcl/Tk support.
If Tkinter is missing, use the client registration audit option or install a portable
Python distribution that includes Tkinter.

## Boot Menu

The build script creates this launcher inside WinPE:

```cmd
X:\PCAudit\start_pc_audit_winpe.cmd
```

On boot, WinPE runs the client registration audit once and then opens a menu:

```text
1. Run PC Audit Tool
2. Run client registration audit once
3. Run full PC Audit app
4. Capture Windows to WIM
5. Apply WIM to disk
6. Command prompt
```

You can also run the menu manually from the WinPE command prompt:

```cmd
X:\PCAudit\start_pc_audit_winpe.cmd
```

## Capture WIM

After Sysprep:

```cmd
C:\Windows\System32\Sysprep\Sysprep.exe /generalize /oobe /shutdown
```

Boot WinPE and run:

```cmd
X:\PCAudit\capture_wim.cmd
```

Default output:

```text
D:\Images\install.wim
```

## Apply WIM

Boot WinPE and run:

```cmd
X:\PCAudit\apply_wim.cmd
```

Warning: applying a WIM can overwrite the target Windows partition. Read the script before using it.

## Add Your Own WIM To The ISO

Copy your Windows image into the toolkit with:

```powershell
.\09_add_wim_image.ps1 -SourceWim "D:\Images\my_windows.wim" -Name "install.wim"
```

Then rebuild:

```powershell
.\01_build_winpe_audit.ps1 -PackageSet Gui
.\03_make_iso.ps1
```

The ISO will contain the image under:

```text
\Images\install.wim
```

Inside WinPE, `X:\PCAudit\apply_wim.cmd` searches the boot media and attached
drives for `\Images\install.wim` or the first `.wim` under `\Images`.

## WDS Import

On a WDS server, run PowerShell as Administrator:

```powershell
.\04_import_to_wds.ps1 -BootWim ".\build\WinPE_PCAudit\media\sources\boot.wim" -InstallWim "D:\Images\install.wim"
```

## Laptop Not Showing in Network Jobs

Make sure:

- The job server is running on the central PC.
- Firewall allows TCP `8765`.
- The client uses the central PC IP, for example:

```text
http://192.168.1.3:8765
```

## Add More Drivers

WinPE injects every `.inf` driver found under:

```text
pxe_wim_toolkit\drivers\
```

You can add a new driver folder with:

```powershell
.\07_add_driver_folder.ps1 -Source "C:\Path\To\Extracted\Drivers" -Name "Dell_Latitude_5420"
```

Then rebuild:

```powershell
.\01_build_winpe_audit.ps1
.\03_make_iso.ps1
```

Use extracted driver folders that contain `.inf` files. If you downloaded a vendor `.exe`,
extract it first, then import the extracted folder.

If DISM reports unsigned drivers and you trust the driver source, rebuild with:

```powershell
.\01_build_winpe_audit.ps1 -PackageSet Gui -ForceUnsignedDrivers
```

If you do not trust the driver, remove that driver folder from `pxe_wim_toolkit\drivers`
and rebuild without `-ForceUnsignedDrivers`.

Recommended driver types for this WinPE:

- LAN / Ethernet
- Wi-Fi, if the model supports WinPE usage
- Storage / NVMe / RAID
- Chipset

## AI Assistance

The PC Audit project, including this toolkit, was developed with the assistance of AI tools.
