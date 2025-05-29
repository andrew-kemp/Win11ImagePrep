Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"
trap { [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)"); exit }

# XAML for Main Window
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Windows 11 Image &amp; USB Creator" Height="630" Width="740" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="1. Select Windows 11 ISO:" VerticalAlignment="Center" Width="170"/>
            <TextBox x:Name="txtISO" Width="400" IsReadOnly="True" Margin="5,0,0,0"/>
            <Button Content="Browse..." x:Name="btnBrowseISO" Margin="5,0,0,0" Width="90"/>
        </StackPanel>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="2. Select Driver MSI file:" VerticalAlignment="Center" Width="170"/>
            <TextBox x:Name="txtMSI" Width="400" IsReadOnly="True" Margin="5,0,0,0"/>
            <Button Content="Browse..." x:Name="btnBrowseMSI" Margin="5,0,0,0" Width="90"/>
        </StackPanel>
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,8">
            <Button Content="Prepare Windows 11 Image with Drivers" x:Name="btnInject" Width="260"/>
            <Button Content="Create Bootable USB" x:Name="btnCreateUSB" Width="230" Margin="10,0,0,0"/>
        </StackPanel>
        <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="3. Select USB Drive:" VerticalAlignment="Center" Width="170"/>
            <ComboBox x:Name="cmbUSB" Width="400"/>
            <Button Content="Refresh USB List" x:Name="btnRefreshUSB" Margin="10,0,0,0" Width="120"/>
        </StackPanel>
        <StackPanel Grid.Row="4" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock x:Name="lblWarn" Foreground="Red" VerticalAlignment="Center" FontWeight="Bold"/>
        </StackPanel>
        <TextBox x:Name="txtLog"
                 Grid.Row="5"
                 Background="Black"
                 Foreground="Lime"
                 FontWeight="Bold"
                 Padding="5,2"
                 FontSize="15"
                 FontFamily="Consolas"
                 HorizontalAlignment="Stretch"
                 VerticalAlignment="Stretch"
                 IsReadOnly="True"
                 BorderThickness="1"
                 BorderBrush="Gray"
                 VerticalScrollBarVisibility="Auto"
                 TextWrapping="Wrap"/>
    </Grid>
</Window>
"@

# --- Helper Paths ---
$topLevelDir = "C:\WinImagePrep"
$windows11Dir = Join-Path $topLevelDir "Windows11"
$driversDir = Join-Path $topLevelDir "Drivers"
$mountDir = Join-Path $topLevelDir "Mount"
$configDir = "C:\WinImagePrep\Config"

# --- Load XAML and Controls ---
$bytes = [System.Text.Encoding]::UTF8.GetBytes($xaml)
$stream = New-Object System.IO.MemoryStream(,$bytes)
$window = [Windows.Markup.XamlReader]::Load($stream)

$txtISO = $window.FindName("txtISO")
$btnBrowseISO = $window.FindName("btnBrowseISO")
$txtMSI = $window.FindName("txtMSI")
$btnBrowseMSI = $window.FindName("btnBrowseMSI")
$btnInject = $window.FindName("btnInject")
$cmbUSB = $window.FindName("cmbUSB")
$btnRefreshUSB = $window.FindName("btnRefreshUSB")
$btnCreateUSB = $window.FindName("btnCreateUSB")
$lblWarn = $window.FindName("lblWarn")
$txtLog = $window.FindName("txtLog")

$script:usbDrives = @()

# --- Logging Helpers ---
function Write-Status([string]$text) {
    $timestamp = Get-Date -Format "HH:mm:ss"
    $txtLog.Dispatcher.Invoke([action]{
        $txtLog.AppendText("[$timestamp] $text`r`n")
        $txtLog.ScrollToEnd()
    })
    $txtLog.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}
function Clear-Log { $txtLog.Dispatcher.Invoke([action]{ $txtLog.Clear() }) }

function Refresh-USBList {
    $cmbUSB.Items.Clear()
    $script:usbDrives = Get-Disk | Where-Object BusType -eq 'USB'
    foreach ($disk in $usbDrives) {
        $desc = "$($disk.Number): $($disk.FriendlyName) - $([math]::Round($disk.Size/1GB,2)) GB"
        $cmbUSB.Items.Add($desc)
    }
    if ($cmbUSB.Items.Count -eq 0) {
        $cmbUSB.Items.Add("No USB drives found")
        $cmbUSB.SelectedIndex = 0
        $btnCreateUSB.IsEnabled = $false
    } else {
        $cmbUSB.SelectedIndex = 0
        if (Test-Path $windows11Dir) {
            $fileCount = (Get-ChildItem $windows11Dir -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
            $btnCreateUSB.IsEnabled = ($fileCount -gt 0)
        }
    }
}
$btnRefreshUSB.Add_Click({ Refresh-USBList })
Refresh-USBList

$cmbUSB.Add_SelectionChanged({
    if ($cmbUSB.SelectedIndex -ge 0 -and $cmbUSB.SelectedItem -notlike "No USB*") {
        $lblWarn.Text = "WARNING: ALL DATA WILL BE ERASED!"
    } else {
        $lblWarn.Text = ""
    }
})

$btnBrowseISO.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "ISO files (*.iso)|*.iso"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtISO.Text = $ofd.FileName
        Write-Status "ISO selected: $($ofd.FileName)"
    }
})

$btnBrowseMSI.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "MSI files (*.msi)|*.msi"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtMSI.Text = $ofd.FileName
        Write-Status "MSI selected: $($ofd.FileName)"
    }
})

function New-BootableWin11USB {
    param(
        [string]$SourceFolder,
        [System.Windows.Controls.ComboBox]$cmbUSB,
        [System.Windows.Controls.Button]$btnCreateUSB,
        [System.Windows.Window]$window
    )
    $btnCreateUSB.IsEnabled = $false
    $window.Cursor = 'Wait'
    Write-Status "===Preparing USB Creation==="
    Write-Status "Starting USB creation..."
    $src = $SourceFolder
    if (-not $src -or -not (Test-Path $src)) {
        Write-Status "Error: No image folder"
        $window.Cursor = 'Arrow'; return
    }
    if ($cmbUSB.SelectedItem -like "No USB*") {
        Write-Status "Error: No USB"
        $window.Cursor = 'Arrow'; return
    }
    $driveNumber = $cmbUSB.SelectedItem.ToString().Split(":")[0]
    $usbDrives = Get-Disk | Where-Object BusType -eq 'USB'
    $usbDisk = $usbDrives | Where-Object Number -eq $driveNumber
    if (-not $usbDisk) {
        Write-Status "Error: Invalid USB disk"
        $window.Cursor = 'Arrow'; return
    }
    if ($usbDisk.Size -lt 14GB) {
        Write-Status "Error: Selected USB drive is less than 14GB!"
        [System.Windows.MessageBox]::Show("Selected USB drive is less than 14GB. Please use a 14GB or larger drive.", "Error", "OK", "Error")
        $window.Cursor = 'Arrow'
        return
    }
    $res = [System.Windows.MessageBox]::Show("WARNING: This will ERASE all partitions and data on USB drive $driveNumber! Continue?", "Confirm", "YesNo", "Warning")
    if ($res -ne "Yes") {
        Write-Status "Aborted"
        $window.Cursor = 'Arrow'; return
    }
    Write-Status "Removing partitions..."
    $existingPartitions = Get-Partition -DiskNumber $driveNumber -ErrorAction SilentlyContinue
    if ($existingPartitions) {
        foreach ($part in $existingPartitions) {
            Remove-Partition -DiskNumber $driveNumber -PartitionNumber $part.PartitionNumber -Confirm:$false -ErrorAction SilentlyContinue
        }
        Start-Sleep -Seconds 2
    }
    $disk = Get-Disk -Number $driveNumber
    if ($disk.PartitionStyle -eq 'RAW') {
        Write-Status "Initializing Disk"
        Initialize-Disk -Number $driveNumber -PartitionStyle MBR
        Start-Sleep -Seconds 2
    }
    $size14GB = 14GB
    Write-Status "Creating 14GB partition..."
    $partition = New-Partition -DiskNumber $driveNumber -Size $size14GB -AssignDriveLetter
    Start-Sleep -Seconds 1

    # --- Get label for USB from config ---
    $usbLabel = "WIN11USB" # default
    $isoLabelFile = Join-Path $configDir "iso-label.txt"
    if (Test-Path $isoLabelFile) {
        $label = Get-Content $isoLabelFile
        if ($label -and $label.Length -le 11) { # FAT32 label max is 11 chars
            $usbLabel = $label
        } elseif ($label) {
            # Truncate and remove illegal chars for FAT32
            $label = $label -replace '[^a-zA-Z0-9_ ]', ''
            $usbLabel = $label.Substring(0, [Math]::Min($label.Length, 11))
        }
    }

    Write-Status "Formatting as FAT32 with label: $usbLabel"
    Format-Volume -Partition $partition -FileSystem FAT32 -NewFileSystemLabel $usbLabel -Confirm:$false
    $usbDriveLetter = ($partition | Get-Volume).DriveLetter
    if (-not $usbDriveLetter) {
        Write-Status "Error: No drive letter"
        $window.Cursor = 'Arrow'; return
    }
    $usbRoot = "$usbDriveLetter`:\"
    Write-Status "Copying files to USB..."
    robocopy "$src" "$usbRoot" /E | Out-Null
    Write-Status "USB creation complete!"
    Write-Status "Bootable Windows 11 USB (14GB FAT32, UEFI compatible) created successfully!"
    $window.Cursor = 'Arrow'
}

$btnInject.Add_Click({
    $iso = $txtISO.Text.Trim()
    $msi = $txtMSI.Text.Trim()
    if (-not (Test-Path $iso)) { Write-Status "Select ISO"; return }
    if (-not (Test-Path $msi)) { Write-Status "Select MSI"; return }
    $window.Cursor = 'Wait'

    # Ensure all required directories exist
    $dirs = @($topLevelDir, $windows11Dir, $driversDir, $mountDir, $configDir)
    foreach ($dir in $dirs) {
        if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    }
    $mountPE = "$mountDir\WinPE"
    $mountSetup = "$mountDir\WinSetup"
    $mountPro = "$mountDir\WinPro"
    $mountEnt = "$mountDir\WinEnt"
    $mountWinRE_Pro = "$mountDir\WinRE_Pro"
    $mountWinRE_Ent = "$mountDir\WinRE_Ent"
    foreach ($d in @($mountPE, $mountSetup, $mountPro, $mountEnt, $mountWinRE_Pro, $mountWinRE_Ent)) {
        if (-not (Test-Path $d)) { New-Item -Path $d -ItemType Directory -Force | Out-Null }
    }
    Write-Status "Working directories created."

    Write-Status "===Mounting ISO==="
    $mountResult = Mount-DiskImage -ImagePath $iso -PassThru
    Start-Sleep -Seconds 2
    $vol = ($mountResult | Get-Volume)
    $driveLetter = $vol.DriveLetter
    if (-not $driveLetter) { Write-Status "ISO mount failed!"; $window.Cursor = 'Arrow'; return }

    # --- Save ISO label to config ---
    $isoLabel = $vol.FileSystemLabel
    if (-not (Test-Path $configDir)) { New-Item -Path $configDir -ItemType Directory -Force | Out-Null }
    Set-Content -Path (Join-Path $configDir "iso-label.txt") -Value $isoLabel

    Write-Status "Copying Files..."
    robocopy "$driveLetter`:\" $windows11Dir /E | Out-Null
    Write-Status "Clearing ReadOnly attributes..."
    Get-ChildItem -Path $windows11Dir -Recurse -File | ForEach-Object { $_.Attributes = $_.Attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly) }
    icacls $windows11Dir /inheritance:e /T | Out-Null
    Dismount-DiskImage -ImagePath $iso

    Write-Status "Extracting Drivers from MSI..."
    if (Test-Path $driversDir) { Remove-Item -Path $driversDir\* -Recurse -Force -ErrorAction SilentlyContinue }
    $msiExtractArgs = "/a `"$msi`" /qb TARGETDIR=`"$driversDir`""
    $proc = Start-Process msiexec -ArgumentList $msiExtractArgs -Wait -PassThru
    if ($proc.ExitCode -ne 0) {
        Write-Status "MSI extraction failed with exit code $($proc.ExitCode)"
        [System.Windows.MessageBox]::Show("Failed to extract MSI. Please check the MSI file.", "Error", "OK", "Error")
        $window.Cursor = 'Arrow'
        return
    }
    # Check for .inf files
    $infFiles = Get-ChildItem -Path $driversDir -Recurse -Filter *.inf -ErrorAction SilentlyContinue
    if (-not $infFiles) {
        Write-Status "No driver (.inf) files found in extracted MSI."
        [System.Windows.MessageBox]::Show("No driver (.inf) files found in the extracted MSI. Ensure this MSI contains drivers, or extract them manually.", "Error", "OK", "Error")
        $window.Cursor = 'Arrow'
        return
    }
    Write-Status "Driver files extracted from MSI."

    $sources = "$windows11Dir\Sources"
    $bootWim = "$sources\boot.wim"
    $installWim = "$sources\install.wim"

    # === Adding drivers to Windows Preinstallation Environment (WinPE) ===
    Write-Status "===Adding Drivers to Windows Preinstallation Environment (WinPE)==="
    Write-Status "Adding drivers to boot.wim index 1 (WinPE)..."
    dism /Mount-Wim /WimFile:"$bootWim" /index:1 /MountDir:"$mountPE" | Out-Null
    dism /Image:"$mountPE" /Add-Driver /Driver:"$driversDir" /Recurse | Out-Null
    dism /Unmount-Wim /MountDir:"$mountPE" /Commit | Out-Null

    # === Adding drivers to Windows Setup (WinSetup) ===
    Write-Status "===Adding Drivers to Windows Setup (WinSetup)==="
    Write-Status "Adding drivers to boot.wim index 2 (WinSetup)..."
    dism /Mount-Wim /WimFile:"$bootWim" /index:2 /MountDir:"$mountSetup" | Out-Null
    dism /Image:"$mountSetup" /Add-Driver /Driver:"$driversDir" /Recurse | Out-Null
    dism /Unmount-Wim /MountDir:"$mountSetup" /Commit | Out-Null

    # === Adding drivers to Windows Professional Edition (Install.wim index 1) ===
    Write-Status "===Adding Drivers to Windows Professional Edition==="
    Write-Status "Adding drivers to install.wim index 1 (Pro)..."
    dism /Mount-Wim /WimFile:"$installWim" /index:1 /MountDir:"$mountPro" | Out-Null
    dism /Image:"$mountPro" /Add-Driver /Driver:"$driversDir" /Recurse | Out-Null

    # === Adding drivers to Windows Recovery Environment (WinRE) for Pro ===
    Write-Status "===Adding Drivers to Windows Recovery Environment (WinRE) for Pro==="
    $winreWim_Pro = Join-Path $mountPro "Windows\System32\Recovery\Winre.wim"
    if (Test-Path $winreWim_Pro) {
        Write-Status "Mounting WinRE for Pro..."
        dism /Mount-Wim /WimFile:"$winreWim_Pro" /index:1 /MountDir:"$mountWinRE_Pro" | Out-Null
        Write-Status "Adding drivers to WinRE for Pro..."
        dism /Image:"$mountWinRE_Pro" /Add-Driver /Driver:"$driversDir" /Recurse | Out-Null
        dism /Unmount-Wim /MountDir:"$mountWinRE_Pro" /Commit | Out-Null
        Write-Status "WinRE driver addition for Pro complete."
        Remove-Item -Path $mountWinRE_Pro -Recurse -Force
    } else {
        Write-Status "No WinRE.wim found for Pro in install.wim index 1. This is normal for most Windows ISOs."
    }
    dism /Unmount-Wim /MountDir:"$mountPro" /Commit | Out-Null

    # === Adding drivers to Windows Enterprise Edition (Install.wim index 2) ===
    Write-Status "===Adding Drivers to Windows Enterprise Edition==="
    Write-Status "Adding drivers to install.wim index 2 (Ent)..."
    dism /Mount-Wim /WimFile:"$installWim" /index:2 /MountDir:"$mountEnt" | Out-Null
    dism /Image:"$mountEnt" /Add-Driver /Driver:"$driversDir" /Recurse | Out-Null

    # === Adding drivers to Windows Recovery Environment (WinRE) for Enterprise ===
    Write-Status "===Adding Drivers to Windows Recovery Environment (WinRE) for Enterprise==="
    $winreWim_Ent = Join-Path $mountEnt "Windows\System32\Recovery\Winre.wim"
    if (Test-Path $winreWim_Ent) {
        Write-Status "Mounting WinRE for Ent..."
        dism /Mount-Wim /WimFile:"$winreWim_Ent" /index:1 /MountDir:"$mountWinRE_Ent" | Out-Null
        Write-Status "Adding drivers to WinRE for Ent..."
        dism /Image:"$mountWinRE_Ent" /Add-Driver /Driver:"$driversDir" /Recurse | Out-Null
        dism /Unmount-Wim /MountDir:"$mountWinRE_Ent" /Commit | Out-Null
        Write-Status "WinRE driver addition for Ent complete."
        Remove-Item -Path $mountWinRE_Ent -Recurse -Force
    } else {
        Write-Status "No WinRE.wim found for Ent in install.wim index 2. This is normal for most Windows ISOs."
    }
    dism /Unmount-Wim /MountDir:"$mountEnt" /Commit | Out-Null

    $wimPath = "$sources\install.wim"
    $swmBasePath = "$sources\install.swm"
    if (Test-Path $wimPath) {
        $wimInfo = Get-Item $wimPath
        if ($wimInfo.Length -gt 4GB) {
            Write-Status "Splitting install.wim"
            Get-ChildItem -Path $sources -Filter "install*.swm" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            $dismCmd = "dism /Split-Image /ImageFile:`"$wimPath`" /SWMFile:`"$swmBasePath`" /FileSize:3800"
            Write-Status "Running: $dismCmd"
            Invoke-Expression $dismCmd
            if (!(Test-Path "$swmBasePath")) {
                Write-Status "Split failed!"
            } else {
                Remove-Item $wimPath -Force
                Write-Status "install.wim split successful. install.swm files created."
            }
        }
    }
    Write-Status "===Image Ready==="
    Write-Status "Driver addition and image preparation complete!"
    $window.Cursor = 'Arrow'
})

$btnCreateUSB.Add_Click({
    New-BootableWin11USB -SourceFolder $windows11Dir -cmbUSB $cmbUSB -btnCreateUSB $btnCreateUSB -window $window
})

$window.ShowDialog() | Out-Null
