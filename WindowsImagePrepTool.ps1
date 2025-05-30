Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"
trap { [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)"); exit }

# --- Custom Quit/Start Again Dialog ---
function Show-CustomChoiceDialog {
    param(
        [string]$Message,
        [string]$Title = "Choose an Option",
        [string]$Button1 = "Quit",
        [string]$Button2 = "Start Again"
    )
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" Height="170" Width="340" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" Topmost="True" WindowStyle="ToolWindow">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <TextBlock Text="$Message" Grid.Row="0" TextWrapping="Wrap" FontSize="15" Margin="0,0,0,12" />
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center">
            <Button x:Name="btn1" Content="$Button1" Width="110" Margin="0,0,12,0" IsDefault="True"/>
            <Button x:Name="btn2" Content="$Button2" Width="110"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    $form = [Windows.Markup.XamlReader]::Load($reader)
    $btn1 = $form.FindName("btn1")
    $btn2 = $form.FindName("btn2")
    $choice = $null
    $btn1.Add_Click({ $choice = 1; $form.Close() })
    $btn2.Add_Click({ $choice = 2; $form.Close() })
    $form.ShowDialog() | Out-Null
    return $choice
}

# --- Saved Image Creation Dialog ---
function Show-CreateFromSavedImageForm {
    $savedImagesDir = "C:\WinImagePrep\SavedImages"
    if (-not (Test-Path $savedImagesDir)) {
        [System.Windows.MessageBox]::Show("No saved images folder found.", "Error", "OK", "Error")
        return
    }
    $folders = Get-ChildItem $savedImagesDir -Directory | Select-Object -ExpandProperty Name
    if (-not $folders) {
        [System.Windows.MessageBox]::Show("No saved images found in $savedImagesDir.", "No Images", "OK", "Warning")
        return
    }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Create USB from Saved Image" Height="470" Width="880" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" Topmost="True" WindowStyle="SingleBorderWindow">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="Select saved image:" VerticalAlignment="Center" Width="180"/>
            <ComboBox x:Name="cmbSaved" Width="510"/>
        </StackPanel>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="Image label:" VerticalAlignment="Center" Width="180"/>
            <TextBlock x:Name="lblImgLabel" Width="510" VerticalAlignment="Center" FontWeight="Bold"/>
        </StackPanel>
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="Select USB Drive:" VerticalAlignment="Center" Width="180"/>
            <ComboBox x:Name="cmbUSB" Width="390"/>
            <Button Content="Refresh USB List" x:Name="btnRefreshUSB" Margin="15,0,0,0" Width="150" Height="24"/>
        </StackPanel>
        <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,0,0,10" HorizontalAlignment="Center">
            <Button Content="Create USB" x:Name="btnCreateUSB" Width="280" Height="24" Margin="0,0,36,0"/>
            <Button Content="Cancel" x:Name="btnCancel" Width="120" Height="24"/>
        </StackPanel>
        <TextBox x:Name="txtSavedLog"
                 Grid.Row="4"
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

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    $form = [Windows.Markup.XamlReader]::Load($reader)
    $cmbSaved = $form.FindName("cmbSaved")
    $lblImgLabel = $form.FindName("lblImgLabel")
    $cmbUSB = $form.FindName("cmbUSB")
    $btnRefreshUSB = $form.FindName("btnRefreshUSB")
    $btnCreateUSB = $form.FindName("btnCreateUSB")
    $btnCancel = $form.FindName("btnCancel")
    $txtSavedLog = $form.FindName("txtSavedLog")

    $folders | ForEach-Object { $cmbSaved.Items.Add($_) }
    $cmbSaved.SelectedIndex = 0

    function Write-SavedStatus([string]$text) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $txtSavedLog.Dispatcher.Invoke([action]{
            $txtSavedLog.AppendText("[$timestamp] $text`r`n")
            $txtSavedLog.ScrollToEnd()
        })
        $txtSavedLog.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    }

    function UpdateLabel {
        $sel = $cmbSaved.SelectedItem
        $imgDir = Join-Path $savedImagesDir $sel
        $lblFile = Join-Path $imgDir "iso-label.txt"
        if (Test-Path $lblFile) {
            $lblImgLabel.Text = Get-Content $lblFile
        } else {
            $lblImgLabel.Text = "(no label found)"
        }
    }
    $cmbSaved.Add_SelectionChanged({ UpdateLabel })
    UpdateLabel

    function RefreshUSBList {
        $cmbUSB.Items.Clear()
        $usbs = Get-Disk | Where-Object BusType -eq 'USB'
        if ($usbs.Count -eq 0) {
            $cmbUSB.Items.Add("No USB drives found")
            $cmbUSB.SelectedIndex = 0
        } else {
            foreach ($disk in $usbs) {
                $desc = "$($disk.Number): $($disk.FriendlyName) - $([math]::Round($disk.Size/1GB,2)) GB"
                $cmbUSB.Items.Add($desc)
            }
            $cmbUSB.SelectedIndex = 0
        }
    }
    $btnRefreshUSB.Add_Click({ RefreshUSBList })
    RefreshUSBList

    $btnCancel.Add_Click({ $form.Close() })

    $btnCreateUSB.Add_Click({
        $sel = $cmbSaved.SelectedItem
        if (-not $sel) { [System.Windows.MessageBox]::Show("Select a saved image.", "Error", "OK", "Error"); return }
        $imgDir = Join-Path $savedImagesDir $sel
        $lblFile = Join-Path $imgDir "iso-label.txt"
        $label = "WIN11USB"
        if (Test-Path $lblFile) { $label = Get-Content $lblFile }
        if ($cmbUSB.SelectedItem -like "No USB*") {
            [System.Windows.MessageBox]::Show("Please insert a USB drive to continue.", "No USB Detected", "OK", "Warning")
            return
        }
        $driveNumber = $cmbUSB.SelectedItem.ToString().Split(":")[0]
        $usbs = Get-Disk | Where-Object BusType -eq 'USB'
        $usbDisk = $usbs | Where-Object Number -eq $driveNumber
        if (-not $usbDisk) {
            [System.Windows.MessageBox]::Show("Invalid USB disk.", "Error", "OK", "Error")
            return
        }
        if ($usbDisk.Size -lt 14GB) {
            [System.Windows.MessageBox]::Show("Selected USB drive is less than 14GB. Please use a 14GB or larger drive.", "Error", "OK", "Error")
            return
        }
        $res = [System.Windows.MessageBox]::Show("WARNING: This will ERASE all partitions and data on USB drive $driveNumber! Continue?", "Confirm", "YesNo", "Warning")
        if ($res -ne "Yes") { return }
        Write-SavedStatus "Removing partitions..."
        $existingPartitions = Get-Partition -DiskNumber $driveNumber -ErrorAction SilentlyContinue
        if ($existingPartitions) {
            foreach ($part in $existingPartitions) {
                Remove-Partition -DiskNumber $driveNumber -PartitionNumber $part.PartitionNumber -Confirm:$false -ErrorAction SilentlyContinue
            }
            Start-Sleep -Seconds 2
        }
        $disk = Get-Disk -Number $driveNumber
        if ($disk.PartitionStyle -eq 'RAW') {
            Write-SavedStatus "Initializing Disk"
            Initialize-Disk -Number $driveNumber -PartitionStyle MBR
            Start-Sleep -Seconds 2
        }
        $size14GB = 14GB
        Write-SavedStatus "Creating 14GB partition..."
        $partition = New-Partition -DiskNumber $driveNumber -Size $size14GB -AssignDriveLetter
        Start-Sleep -Seconds 1
        if ($label.Length -gt 11) { $label = $label.Substring(0, 11) }
        $label = $label -replace '[^a-zA-Z0-9_ ]', ''
        Write-SavedStatus "Formatting as FAT32 with label: $label"
        Format-Volume -Partition $partition -FileSystem FAT32 -NewFileSystemLabel $label -Confirm:$false
        $usbDriveLetter = ($partition | Get-Volume).DriveLetter
        if (-not $usbDriveLetter) {
            [System.Windows.MessageBox]::Show("Could not get drive letter for USB.", "Error", "OK", "Error")
            return
        }
        $usbRoot = "$usbDriveLetter`:\"
        Write-SavedStatus "Copying files to USB..."
        robocopy "$imgDir" "$usbRoot" /E | Out-Null
        Write-SavedStatus "USB creation complete!"
        [System.Windows.MessageBox]::Show("Bootable Windows 11 USB created successfully from saved image!", "Done", "OK", "Info")
        $form.Close()
    })

    $form.ShowDialog() | Out-Null
}

# --- USB from ISO Dialog (with label and temp cleanup) ---
function Show-UsbFromIsoForm {
    $isoTempDir = "C:\WinImagePrep\ISO_Temp"
    if (-not (Test-Path $isoTempDir)) { New-Item -Path $isoTempDir -ItemType Directory -Force | Out-Null }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Create Bootable USB from ISO" Height="360" Width="760" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" Topmost="True" WindowStyle="SingleBorderWindow">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="Select ISO:" VerticalAlignment="Center" Width="120"/>
            <TextBox x:Name="txtIsoOnly" Width="400" IsReadOnly="True"/>
            <Button Content="Browse..." x:Name="btnBrowseIsoOnly" Width="90" Height="24" Margin="10,0,0,0"/>
        </StackPanel>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBlock Text="Select USB Drive:" VerticalAlignment="Center" Width="120"/>
            <ComboBox x:Name="cmbUsbIsoOnly" Width="350"/>
            <Button Content="Refresh" x:Name="btnRefreshUsbIsoOnly" Width="110" Height="24" Margin="10,0,0,0"/>
        </StackPanel>
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,10" HorizontalAlignment="Center">
            <Button Content="Create USB from ISO" x:Name="btnCreateUsbIsoOnly" Width="260" Height="24" IsEnabled="False"/>
            <Button Content="Cancel" x:Name="btnCancelUsbIsoOnly" Width="120" Height="24" Margin="20,0,0,0"/>
        </StackPanel>
        <TextBox x:Name="txtLogIsoOnly"
                 Grid.Row="3"
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

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    $form = [Windows.Markup.XamlReader]::Load($reader)
    $txtIsoOnly = $form.FindName("txtIsoOnly")
    $btnBrowseIsoOnly = $form.FindName("btnBrowseIsoOnly")
    $cmbUsbIsoOnly = $form.FindName("cmbUsbIsoOnly")
    $btnRefreshUsbIsoOnly = $form.FindName("btnRefreshUsbIsoOnly")
    $btnCreateUsbIsoOnly = $form.FindName("btnCreateUsbIsoOnly")
    $btnCancelUsbIsoOnly = $form.FindName("btnCancelUsbIsoOnly")
    $txtLogIsoOnly = $form.FindName("txtLogIsoOnly")

    function Write-Log([string]$text) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $txtLogIsoOnly.Dispatcher.Invoke([action]{
            $txtLogIsoOnly.AppendText("[$timestamp] $text`r`n")
            $txtLogIsoOnly.ScrollToEnd()
        })
        $txtLogIsoOnly.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    }

    function RefreshUsbList {
        $cmbUsbIsoOnly.Items.Clear()
        $usbs = Get-Disk | Where-Object BusType -eq 'USB'
        if ($usbs.Count -eq 0) {
            $cmbUsbIsoOnly.Items.Add("No USB drives found")
            $cmbUsbIsoOnly.SelectedIndex = 0
        } else {
            foreach ($disk in $usbs) {
                $desc = "$($disk.Number): $($disk.FriendlyName) - $([math]::Round($disk.Size/1GB,2)) GB"
                $cmbUsbIsoOnly.Items.Add($desc)
            }
            $cmbUsbIsoOnly.SelectedIndex = 0
        }
    }

    $btnRefreshUsbIsoOnly.Add_Click({ RefreshUsbList })
    RefreshUsbList

    $btnCancelUsbIsoOnly.Add_Click({ $form.Close() })

    $btnBrowseIsoOnly.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "ISO files (*.iso)|*.iso"
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtIsoOnly.Text = $ofd.FileName
            Write-Log "ISO selected: $($ofd.FileName)"
            if ($cmbUsbIsoOnly.SelectedIndex -ge 0 -and $cmbUsbIsoOnly.SelectedItem -notlike "No USB*") {
                $btnCreateUsbIsoOnly.IsEnabled = $true
            }
        }
    })

    $cmbUsbIsoOnly.Add_SelectionChanged({
        if ($cmbUsbIsoOnly.SelectedIndex -ge 0 -and $cmbUsbIsoOnly.SelectedItem -notlike "No USB*" -and $txtIsoOnly.Text.Length -gt 0) {
            $btnCreateUsbIsoOnly.IsEnabled = $true
        } else {
            $btnCreateUsbIsoOnly.IsEnabled = $false
        }
    })

    $btnCreateUsbIsoOnly.Add_Click({
        $isoPath = $txtIsoOnly.Text.Trim()
        if (-not (Test-Path $isoPath)) {
            Write-Log "Select a valid ISO file."
            return
        }
        if ($cmbUsbIsoOnly.SelectedItem -like "No USB*") {
            Write-Log "No USB selected."
            return
        }
        $usbDriveNumber = $cmbUsbIsoOnly.SelectedItem.ToString().Split(":")[0]
        $usbs = Get-Disk | Where-Object BusType -eq 'USB'
        $usbDisk = $usbs | Where-Object Number -eq $usbDriveNumber
        if (-not $usbDisk) {
            Write-Log "Invalid USB selection."
            return
        }
        $usbSizeGB = [math]::Round($usbDisk.Size/1GB,2)
        $res = [System.Windows.MessageBox]::Show("WARNING: This will ERASE all partitions and data on USB drive $usbDriveNumber! Continue?", "Confirm", "YesNo", "Warning")
        if ($res -ne "Yes") { Write-Log "Operation cancelled by user."; return }

        Write-Log "Clearing temp folder..."
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "$isoTempDir\*" | Out-Null

        Write-Log "Mounting ISO..."
        $mountResult = Mount-DiskImage -ImagePath $isoPath -PassThru
        Start-Sleep -Seconds 2
        $vol = ($mountResult | Get-Volume)
        $driveLetter = $vol.DriveLetter
        $isoLabel = $vol.FileSystemLabel
        if (-not $driveLetter) { Write-Log "ISO mount failed!"; return }

        Write-Log "Copying ISO files to temp..."
        robocopy "$driveLetter`:\" $isoTempDir /E | Out-Null

        Write-Log "Clearing ReadOnly attributes..."
        Get-ChildItem -Path $isoTempDir -Recurse -File | ForEach-Object { $_.Attributes = $_.Attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly) }
        icacls $isoTempDir /inheritance:e /T | Out-Null
        Dismount-DiskImage -ImagePath $isoPath

        $sources = "$isoTempDir\Sources"
        $wimPath = "$sources\install.wim"
        $swmBasePath = "$sources\install.swm"
        if (Test-Path $wimPath) {
            $wimInfo = Get-Item $wimPath
            if ($wimInfo.Length -gt 4GB) {
                Write-Log "install.wim > 4GB, splitting..."
                Get-ChildItem -Path $sources -Filter "install*.swm" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
                & dism /Split-Image /ImageFile:"$wimPath" /SWMFile:"$swmBasePath" /FileSize:3800 | Out-Null
                if (!(Test-Path "$swmBasePath")) {
                    Write-Log "Split failed!"
                } else {
                    Remove-Item $wimPath -Force
                    Write-Log "install.wim split successful. install.swm files created."
                }
            }
        }

        if ($usbSizeGB -lt 32) {
            $partitionSizeGB = [math]::Floor($usbSizeGB)
        } else {
            $partitionSizeGB = 32
        }
        Write-Log "Preparing USB: $usbSizeGB GB detected, partition size: $partitionSizeGB GB"
        $existingPartitions = Get-Partition -DiskNumber $usbDriveNumber -ErrorAction SilentlyContinue
        if ($existingPartitions) {
            foreach ($part in $existingPartitions) {
                Remove-Partition -DiskNumber $usbDriveNumber -PartitionNumber $part.PartitionNumber -Confirm:$false -ErrorAction SilentlyContinue
            }
            Start-Sleep -Seconds 2
        }
        $disk = Get-Disk -Number $usbDriveNumber
        if ($disk.PartitionStyle -eq 'RAW') {
            Write-Log "Initializing Disk"
            Initialize-Disk -Number $usbDriveNumber -PartitionStyle MBR
            Start-Sleep -Seconds 2
        }
        Write-Log "Creating partition..."
        $partition = New-Partition -DiskNumber $usbDriveNumber -Size ($partitionSizeGB * 1GB) -AssignDriveLetter
        Start-Sleep -Seconds 1

        $label = $isoLabel
        if ($null -eq $label -or $label -eq "") { $label = "WIN11USB" }
        $label = $label -replace '[^a-zA-Z0-9_ ]', ''
        if ($label.Length -gt 11) { $label = $label.Substring(0,11) }

        Write-Log "Formatting as FAT32 with label: $label"
        Format-Volume -Partition $partition -FileSystem FAT32 -NewFileSystemLabel $label -Confirm:$false
        $usbDriveLetter = ($partition | Get-Volume).DriveLetter
        if (-not $usbDriveLetter) {
            Write-Log "Error: No drive letter"
            return
        }
        $usbRoot = "$usbDriveLetter`:\"
        Write-Log "Copying files to USB..."
        robocopy "$isoTempDir" "$usbRoot" /E | Out-Null
        Write-Log "USB creation complete!"

        Write-Log "Cleaning up temp ISO files..."
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue "$isoTempDir\*" | Out-Null

        [System.Windows.MessageBox]::Show("Bootable Windows 11 USB created successfully from ISO!", "Done", "OK", "Info")
        $form.Close()
    })

    $form.ShowDialog() | Out-Null
}

# --- Main Window ---
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Windows 11 Image &amp; USB Creator" Height="655" Width="860" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
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
            <Button Content="Browse..." x:Name="btnBrowseISO" Margin="5,0,0,0" Width="90" Height="24"/>
        </StackPanel>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="2. Select Driver MSI file:" VerticalAlignment="Center" Width="170"/>
            <TextBox x:Name="txtMSI" Width="400" IsReadOnly="True" Margin="5,0,0,0"/>
            <Button Content="Browse..." x:Name="btnBrowseMSI" Margin="5,0,0,0" Width="90" Height="24"/>
        </StackPanel>
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,8">
            <Button Content="Prepare Windows 11 Image with Drivers" x:Name="btnInject" Width="260" Height="24"/>
            <Button Content="Create from Saved Image" x:Name="btnFromSaved" Width="245" Margin="10,0,0,0" Height="24"/>
            <Button Content="Create Bootable USB from ISO" x:Name="btnUsbFromIso" Width="280" Height="24" Margin="10,0,0,0"/>
        </StackPanel>
        <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="3. Select USB Drive:" VerticalAlignment="Center" Width="170"/>
            <ComboBox x:Name="cmbUSB" Width="400"/>
            <Button Content="Refresh USB List" x:Name="btnRefreshUSB" Margin="10,0,0,0" Width="120" Height="24"/>
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

$topLevelDir   = "C:\WinImagePrep"
$windows11Dir  = Join-Path $topLevelDir "Windows11"
$driversDir    = Join-Path $topLevelDir "Drivers"
$mountDir      = Join-Path $topLevelDir "Mount"
$configDir     = "C:\WinImagePrep\Config"

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
$btnFromSaved = $window.FindName("btnFromSaved")
$btnUsbFromIso = $window.FindName("btnUsbFromIso")
$lblWarn = $window.FindName("lblWarn")
$txtLog = $window.FindName("txtLog")
$script:usbDrives = @()

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
    if ($script:usbDrives.Count -eq 0) {
        $cmbUSB.Items.Add("No USB drives found")
        $cmbUSB.SelectedIndex = 0
    } else {
        foreach ($disk in $usbDrives) {
            $desc = "$($disk.Number): $($disk.FriendlyName) - $([math]::Round($disk.Size/1GB,2)) GB"
            $cmbUSB.Items.Add($desc)
        }
        $cmbUSB.SelectedIndex = 0
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

$btnFromSaved.Add_Click({ Show-CreateFromSavedImageForm })
$btnUsbFromIso.Add_Click({ Show-UsbFromIsoForm })

function New-BootableWin11USB {
    param(
        [string]$SourceFolder,
        [System.Windows.Controls.ComboBox]$cmbUSB,
        [object]$btnCreateUSB,
        [System.Windows.Window]$window
    )
    if ($script:usbDrives.Count -eq 0) {
        Write-Status "No USB drives found. Please insert a USB drive."
        [System.Windows.MessageBox]::Show("Please insert a USB drive to continue.", "No USB Detected", "OK", "Warning")
        $window.Cursor = 'Arrow'
        return
    }
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
        [System.Windows.MessageBox]::Show("Please insert a USB drive to continue.", "No USB Detected", "OK", "Warning")
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
    $usbLabel = "WIN11USB"
    $isoLabelFile = Join-Path $configDir "iso-label.txt"
    if (Test-Path $isoLabelFile) {
        $label = Get-Content $isoLabelFile
        if ($label -and $label.Length -le 11) {
            $usbLabel = $label
        } elseif ($label) {
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

    $savePrompt = [System.Windows.MessageBox]::Show(
        "Do you want to save the prepared Windows 11 image for later use?", 
        "Save Prepared Image", 
        [System.Windows.MessageBoxButton]::YesNo, 
        [System.Windows.MessageBoxImage]::Question
    )
    if ($savePrompt -eq [System.Windows.MessageBoxResult]::Yes) {
        $savedImagesDir = "C:\WinImagePrep\SavedImages"
        if (-not (Test-Path $savedImagesDir)) {
            New-Item -Path $savedImagesDir -ItemType Directory -Force | Out-Null
        }
        Add-Type -AssemblyName Microsoft.VisualBasic
        $customName = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter a name for the saved image folder (or leave blank for timestamp):", 
            "Save Prepared Image", 
            ""
        )
        if (-not $customName -or $customName.Trim() -eq "") {
            $customName = "Image_" + (Get-Date -Format "yyyyMMdd_HHmmss")
        } else {
            $invalid = [System.IO.Path]::GetInvalidFileNameChars()
            foreach ($char in $invalid) {
                $customName = $customName -replace [Regex]::Escape([string]$char), ""
            }
        }
        $saveSubdir = Join-Path $savedImagesDir $customName
        Write-Status "Saving prepared image to $saveSubdir..."
        robocopy "$windows11Dir" "$saveSubdir" /E | Out-Null
        Write-Status "Saved prepared files to $saveSubdir."
        $isoLabelFile = Join-Path $configDir "iso-label.txt"
        if (Test-Path $isoLabelFile) {
            Copy-Item $isoLabelFile -Destination $saveSubdir -Force
            Remove-Item $isoLabelFile -Force
            Write-Status "Copied image label to $saveSubdir and cleaned up config."
        }
    } else {
        Write-Status "Prepared files discarded by user choice."
        $isoLabelFile = Join-Path $configDir "iso-label.txt"
        if (Test-Path $isoLabelFile) {
            Remove-Item $isoLabelFile -Force
        }
    }
    Write-Status "Cleaning up working folders..."
    if (Test-Path $windows11Dir) {
        Get-ChildItem -Path $windows11Dir -Recurse -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path $driversDir) {
        Get-ChildItem -Path $driversDir -Recurse -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
    Write-Status "Cleanup complete."

    $againPrompt = Show-CustomChoiceDialog "Would you like to quit or start again?" "All Done" "Quit" "Start Again"
    if ($againPrompt -eq 1) {
        exit
    } else {
        $txtISO.Clear()
        $txtMSI.Clear()
        $cmbUSB.SelectedIndex = -1
        $lblWarn.Text = ""
        Clear-Log
        Refresh-USBList
    }
}

$btnInject.Add_Click({
    $iso = $txtISO.Text.Trim()
    $msi = $txtMSI.Text.Trim()
    if (-not (Test-Path $iso)) { Write-Status "Select ISO"; return }
    if (-not (Test-Path $msi)) { Write-Status "Select MSI"; return }
    $window.Cursor = 'Wait'

    function DismStep {
        param(
            [string]$Description,
            [scriptblock]$RealCommand
        )
        Write-Status $Description
        Invoke-Command $RealCommand
    }

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

    Write-Status "===Adding Drivers to Windows Preinstallation Environment (WinPE)==="
    DismStep "Adding drivers to boot.wim index 1 (WinPE)..." {
        dism /Mount-Wim /WimFile:"$bootWim" /index:1 /MountDir:"$mountPE" | Out-Null
        dism /Image:"$mountPE" /Add-Driver /Driver:"$driversDir" /Recurse | Out-Null
        dism /Unmount-Wim /MountDir:"$mountPE" /Commit | Out-Null
    }

    Write-Status "===Adding Drivers to Windows Setup (WinSetup)==="
    DismStep "Adding drivers to boot.wim index 2 (WinSetup)..." {
        dism /Mount-Wim /WimFile:"$bootWim" /index:2 /MountDir:"$mountSetup" | Out-Null
        dism /Image:"$mountSetup" /Add-Driver /Driver:"$driversDir" /Recurse | Out-Null
        dism /Unmount-Wim /MountDir:"$mountSetup" /Commit | Out-Null
    }

    Write-Status "===Adding Drivers to Windows Professional Edition==="
    DismStep "Adding drivers to install.wim index 1 (Pro)..." {
        dism /Mount-Wim /WimFile:"$installWim" /index:1 /MountDir:"$mountPro" | Out-Null
        dism /Image:"$mountPro" /Add-Driver /Driver:"$driversDir" /Recurse | Out-Null
    }

    Write-Status "===Adding Drivers to Windows Recovery Environment (WinRE) for Pro==="
    $winreWim_Pro = Join-Path $mountPro "Windows\System32\Recovery\Winre.wim"
    if (Test-Path $winreWim_Pro) {
        Write-Status "Mounting WinRE for Pro..."
        DismStep "Adding drivers to WinRE for Pro..." {
            dism /Mount-Wim /WimFile:"$winreWim_Pro" /index:1 /MountDir:"$mountWinRE_Pro" | Out-Null
            dism /Image:"$mountWinRE_Pro" /Add-Driver /Driver:"$driversDir" /Recurse | Out-Null
            dism /Unmount-Wim /MountDir:"$mountWinRE_Pro" /Commit | Out-Null
            Remove-Item -Path $mountWinRE_Pro -Recurse -Force
        }
        Write-Status "WinRE driver addition for Pro complete."
    } else {
        Write-Status "No WinRE.wim found for Pro in install.wim index 1. This is normal for most Windows ISOs."
    }
    DismStep "Unmounting install.wim index 1 (Pro)..." {
        dism /Unmount-Wim /MountDir:"$mountPro" /Commit | Out-Null
    }

    Write-Status "===Adding Drivers to Windows Enterprise Edition==="
    DismStep "Adding drivers to install.wim index 2 (Ent)..." {
        dism /Mount-Wim /WimFile:"$installWim" /index:2 /MountDir:"$mountEnt" | Out-Null
        dism /Image:"$mountEnt" /Add-Driver /Driver:"$driversDir" /Recurse | Out-Null
    }

    Write-Status "===Adding Drivers to Windows Recovery Environment (WinRE) for Enterprise==="
    $winreWim_Ent = Join-Path $mountEnt "Windows\System32\Recovery\Winre.wim"
    if (Test-Path $winreWim_Ent) {
        Write-Status "Mounting WinRE for Ent..."
        DismStep "Adding drivers to WinRE for Ent..." {
            dism /Mount-Wim /WimFile:"$winreWim_Ent" /index:1 /MountDir:"$mountWinRE_Ent" | Out-Null
            dism /Image:"$mountWinRE_Ent" /Add-Driver /Driver:"$driversDir" /Recurse | Out-Null
            dism /Unmount-Wim /MountDir:"$mountWinRE_Ent" /Commit | Out-Null
            Remove-Item -Path $mountWinRE_Ent -Recurse -Force
        }
        Write-Status "WinRE driver addition for Ent complete."
    } else {
        Write-Status "No WinRE.wim found for Ent in install.wim index 2. This is normal for most Windows ISOs."
    }
    DismStep "Unmounting install.wim index 2 (Ent)..." {
        dism /Unmount-Wim /MountDir:"$mountEnt" /Commit | Out-Null
    }

    $wimPath = "$sources\install.wim"
    $swmBasePath = "$sources\install.swm"
    if (Test-Path $wimPath) {
        $wimInfo = Get-Item $wimPath
        if ($wimInfo.Length -gt 4GB) {
            Write-Status "Splitting install.wim"
            Get-ChildItem -Path $sources -Filter "install*.swm" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            DismStep "Splitting install.wim..." {
                & dism /Split-Image /ImageFile:"$wimPath" /SWMFile:"$swmBasePath" /FileSize:3800
            }
            if (!(Test-Path "$swmBasePath")) {
                Write-Status "Split failed!"
            } else {
                Remove-Item $wimPath -Force
                Write-Status "install.wim split successful. install.swm files created."
            }
        }
    }
    Write-Status "===Image Ready==="
    Write-Status "Driver addition and image preparation complete."

    do {
        $script:usbDrives = Get-Disk | Where-Object BusType -eq 'USB'
        Refresh-USBList
        $usbPresent = ($script:usbDrives | Measure-Object).Count -gt 0
        $driveList = $script:usbDrives | ForEach-Object { "$($_.Number): $($_.FriendlyName) - $([math]::Round($_.Size/1GB,1)) GB [Status: $($_.OperationalStatus)]" }
        $driveMsg = if ($driveList) { "`n`nDetected USBs:`n" + ($driveList -join "`n") } else { "" }
        if (-not $usbPresent) {
            [System.Windows.MessageBox]::Show("Please insert a USB drive to continue.$driveMsg", "Insert USB", "OK", "Warning")
            Start-Sleep -Seconds 3
        }
    } while (-not $usbPresent)
    Start-Sleep -Seconds 2
    Refresh-USBList

    $go = [System.Windows.MessageBox]::Show("Image preparation complete. You're now ready to create the Bootable USB. Click OK to continue.", "Ready to Create USB", "OK", "Information")
    if ($go -eq "OK") {
        if ($cmbUSB.SelectedIndex -lt 0 -or $cmbUSB.SelectedItem -like "No USB*") {
            $cmbUSB.SelectedIndex = 0
        }
        New-BootableWin11USB -SourceFolder $windows11Dir -cmbUSB $cmbUSB -btnCreateUSB $null -window $window
        return
    }
})

$window.ShowDialog() | Out-Null
