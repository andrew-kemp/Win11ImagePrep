# Windows 11 Image & USB Creator

A PowerShell GUI tool for preparing Windows 11 installation images with injected drivers and creating bootable USB drives (UEFI-compatible, FAT32, 14GB+). This script is especially useful for creating Windows 11 USB installers with additional drivers (from MSI packages), and for managing saved/prepared images.

## Features

- **Graphical User Interface (GUI)** using WPF/XAML (no need for manual command-line work)
- Mount and extract Windows 11 ISOs
- Extract and inject drivers from MSI packages into the Windows image (boot.wim/index 1 & 2, install.wim/index 1 & 2, and WinRE.wim if present)
- Automatically handles >4GB `install.wim` by splitting to `install.swm` for FAT32 compatibility
- Create bootable USB drives from prepared image folders or directly from ISO
- Save and reuse custom-prepared Windows images
- All destructive operations prompt the user for confirmation
- Robust error handling with user-friendly message boxes

## Requirements

- **Windows 10/11** (must be run as Administrator)
- **PowerShell 5.1+**
- **.NET Framework 4.6+**
- **DISM** (Deployment Image Servicing and Management) command-line tool (included in Windows)
- **Robocopy** (included in Windows)
- **msiexec** (included in Windows)
- **ISO and USB drive (14GB+)**
- **Drivers in MSI format**

## Usage

1. **Download or copy the script** to your local machine (e.g. `WinImagePrep.ps1`).

2. **Run PowerShell as Administrator.**

3. **Execute the script**:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\WindowsImagePrepTool.ps1
   ```

4. **Follow the GUI prompts:**
    - **Select a Windows 11 ISO** file.
    - **Select a driver MSI** file (optional, but recommended for OEM drivers).
    - **Click "Prepare Windows 11 Image with Drivers"** to inject drivers into the Windows image.
    - When prompted, **insert a USB drive** (14GB or larger; all data will be erased), and select it from the list.
    - The script will format the USB and copy files, creating a UEFI bootable Windows 11 USB.
    - Optionally, **save the prepared image** for future use.
    - **Alternative creation options:**
        - "Create from Saved Image": Use a previously prepared and saved Windows image folder.
        - "Create Bootable USB from ISO": Write a USB directly from an ISO (with >4GB `install.wim` handled).

## Saved Images

Prepared images can be saved (under `C:\WinImagePrep\SavedImages`) for quick future USB creation without re-injecting drivers or re-processing ISOs.

## Important Notes

- **All data on the selected USB drive will be erased.** Triple-check the drive selection.
- For ISOs with `install.wim` larger than 4GB, the script splits it to multiple `install.swm` files for FAT32 compatibility.
- Some operations (mounting, formatting, partitioning) can take several minutes.
- This script is designed for clean, UEFI-compatible USB creation and driver integration.

## Troubleshooting

- **Run as Administrator**: Most failures are due to insufficient permissions.
- **Check for required tools**: `DISM`, `robocopy`, and `msiexec` must be available.
- **Driver MSI must actually contain drivers**: Not all MSI files are suitable.
- **All actions are logged in the GUI**: Check the log window for errors or status.

## Folder Structure

- `C:\WinImagePrep\Windows11` — Working directory for Windows files
- `C:\WinImagePrep\Drivers` — Extracted drivers from MSI
- `C:\WinImagePrep\SavedImages` — User-saved prepared images
- `C:\WinImagePrep\Config` — Temporary config (e.g., ISO label)
- `C:\WinImagePrep\ISO_Temp` — Temporary ISO extraction
- `C:\WinImagePrep\Mount` — Temporary Mount root folder for WIM files

## Sample Workflow

1. **Prepare USB with drivers:**
    - Select ISO and MSI
    - Prepare image
    - Insert USB, confirm prompts
    - (Optionally) save prepared image

2. **Quick-create USB from previously saved image:**
    - Click "Create from Saved Image"
    - Choose image and USB, create

3. **Direct USB creation from ISO (no driver injection):**
    - Click "Create Bootable USB from ISO"
    - Select ISO and USB, create

## Known Limitations

- Only supports x64 and ARM Windows 11 ISOs and standard install.wim layout.
- Requires at least 14GB USB drive.
- Not designed for legacy BIOS (FAT32/UEFI only).
- Some anti-virus or endpoint protection software may interfere with file operations.

## Security & Safety

- The script does **not** collect or transmit any data.
- All file operations are local.
- All destructive actions require explicit user confirmation.

## License

MIT License (or as specified by repository owner).

---

**Author:**  
andrew-kemp

**Contributions and issues welcome!**

```
