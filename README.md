# Windows 11 Image & USB Creator

## Synopsis

This PowerShell script provides a **graphical user interface (GUI)** for preparing a Windows 11 installation image with custom drivers and creating a bootable USB drive. It is intended for IT pros or advanced users who need to inject drivers into a Windows 11 ISO and then create a UEFI-compatible USB installer.

### Key Features

- **GUI Built with XAML & WPF**  
  Uses XAML for the window layout and PowerShell to bind logic, creating a modern, user-friendly interface.

- **Step-by-Step Workflow**
    1. **Select a Windows 11 ISO file.**
    2. **Select a driver MSI file** (which will be extracted to obtain `.inf` drivers).
    3. **Prepare and inject drivers into the Windows image.**
    4. **Select a USB drive and create a bootable USB installer.**

- **Driver Injection**
    - Mounts the selected Windows 11 ISO.
    - Copies the contents to a working directory.
    - Extracts drivers from the selected MSI package.
    - Uses DISM to inject those drivers into various Windows image indexes:
      - `boot.wim` (WinPE and Setup)
      - `install.wim` (Pro and Enterprise editions, and their WinRE environments)
    - Supports splitting `install.wim` into `.swm` files if it exceeds 4GB (for FAT32 compatibility).

- **Bootable USB Creation**
    - Detects all USB drives and allows the user to select one.
    - Warns that **all data on the selected USB will be erased**.
    - Partitions and formats the USB as FAT32 (14GB partition, UEFI bootable).
    - Copies the prepared Windows image files to the USB.

- **Logging & Error Handling**
    - All actions are logged in a scrolling text box with timestamps.
    - Errors are caught and displayed in a message box, and logged.
    - The app disables/enables buttons as appropriate to prevent invalid operations.

- **Helper Functions**
    - Refresh USB list, browse for files, and clear log.
    - Ensures all necessary working directories exist.

### Technical Notes

- Uses PowerShell cmdlets (`Mount-DiskImage`, `Get-Disk`, `Remove-Partition`, `Format-Volume`, `robocopy`, `dism`, etc.).
- Handles both GUI events and backend logic in the same script.
- Designed for use on Windows systems with necessary permissions and tools (DISM, robocopy, etc.).
- The script is not intended for general users—administrative privileges are generally required for disk/partition actions.

---

**In summary:**  
This script provides a full, guided experience for injecting drivers into a Windows 11 ISO and making a bootable USB installer—with all steps, logging, and error handling managed via a GUI.
