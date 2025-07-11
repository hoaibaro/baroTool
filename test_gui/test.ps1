# ADMIN PRIVILEGES CHECK & INITIALIZATION
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "This script requires administrative privileges. Attempting to restart with elevation..."

    # Restart script with admin privileges
    $scriptPath = $MyInvocation.MyCommand.Path
    $arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""

    Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs

    # Exit the current non-elevated instance
    exit
}

# ẨN CONSOLE
try {
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    ' -ErrorAction SilentlyContinue

    $consolePtr = [Console.Window]::GetConsoleWindow()
    if ($consolePtr -ne [System.IntPtr]::Zero) {
        [Console.Window]::ShowWindow($consolePtr, 0)
    }
}
catch {
}

# Ẩn menu chính
function Hide-MainMenu {
    $script:form.Hide()
}

# Hiện menu chính
function Show-MainMenu {
    $script:form.Show()
}

# Tạo nút động
function New-DynamicButton {
    param (
        [string]$text,
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height,
        [scriptblock]$clickAction,
        [System.Drawing.Color]$normalColor = [System.Drawing.Color]::FromArgb(0, 128, 0),
        [System.Drawing.Color]$hoverColor = [System.Drawing.Color]::FromArgb(0, 180, 0),
        [System.Drawing.Color]$pressColor = [System.Drawing.Color]::FromArgb(0, 100, 0),
        [System.Drawing.Color]$textColor = [System.Drawing.Color]::White,
        [string]$fontName = "Arial",
        [int]$fontSize = 12,
        [System.Drawing.FontStyle]$fontStyle = [System.Drawing.FontStyle]::Bold
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $text
    $button.Location = New-Object System.Drawing.Point($x, $y)
    $button.Size = New-Object System.Drawing.Size($width, $height)
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.BackColor = $normalColor
    $button.ForeColor = $textColor
    $button.Font = New-Object System.Drawing.Font($fontName, $fontSize, $fontStyle)
    $button.FlatAppearance.BorderSize = 0
    $button.FlatAppearance.MouseOverBackColor = $hoverColor
    $button.FlatAppearance.MouseDownBackColor = $pressColor
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $button.Add_Click($clickAction)

    return $button
}

# Lệnh kiểm tra và tải thư viện Windows Forms
Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
Add-Type -AssemblyName System.Drawing -ErrorAction Stop

# Tạo form chính có thể thay đổi kích thước
$script:form = New-Object System.Windows.Forms.Form
$script:form.Text = "BAOPROVIP - SYSTEM MANAGEMENT"
$script:form.Size = New-Object System.Drawing.Size(600, 400)
$script:form.MinimumSize = New-Object System.Drawing.Size(600, 400)  # Kích thước tối thiểu
$script:form.StartPosition = "CenterScreen"
$script:form.BackColor = [System.Drawing.Color]::Black
$script:form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable  # CHỖ NÀY THAY ĐỔI
$script:form.MaximizeBox = $true  # Cho phép maximize

# Thêm màu gradient với resize handling
$script:form.Add_Paint({
        $graphics = $_.Graphics
        $rect = New-Object System.Drawing.Rectangle(0, 0, $script:form.Width, $script:form.Height)
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $rect,
            [System.Drawing.Color]::FromArgb(0, 0, 0),
            [System.Drawing.Color]::FromArgb(0, 30, 0),
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
        )
        $graphics.FillRectangle($brush, $rect)
        $brush.Dispose()
    })

# Tiêu đề - RESPONSIVE
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "WELCOME TO BAOPROVIP"
$titleLabel.Font = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$titleLabel.Size = New-Object System.Drawing.Size($script:form.ClientSize.Width, 60)
$titleLabel.Location = New-Object System.Drawing.Point(0, 20)
$titleLabel.BackColor = [System.Drawing.Color]::Transparent
$titleLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:form.Controls.Add($titleLabel)


# Hàm thông báo trạng thái
function Add-Status {
    param(
        [string]$message,
        [System.Windows.Forms.TextBox]$statusTextBox
    )

    if ($statusTextBox.Text -eq "Please select a device type...") {
        $statusTextBox.Clear()
    }

    $timestamp = Get-Date -Format "HH:mm:ss"
    $statusTextBox.AppendText("[$timestamp] $message`r`n")
    $statusTextBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# [4] Volume Management Functions
function Update-DriveList {
    $driveListBox.Items.Clear()
    $drives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } }, `
    @{Name = 'VolumeName'; Expression = { $_.VolumeName } }, `
    @{Name = 'Size (GB)'; Expression = { [math]::round($_.Size / 1GB, 0) } }, `
    @{Name = 'FreeSpace (GB)'; Expression = { [math]::round($_.FreeSpace / 1GB, 0) } }

    foreach ($drive in $drives) {
        $driveInfo = "$($drive.Name) - $($drive.VolumeName) - Size: $($drive.'Size (GB)') GB - Free: $($drive.'FreeSpace (GB)') GB"
        $driveListBox.Items.Add($driveInfo)
    }

    if ($driveListBox.Items.Count -gt 0) {
        $driveListBox.SelectedIndex = 0
    }

    return $drives.Count
}

# [4.3] RENAME VOLUME FUNCTIONS
function New-RenameVolumeTitle {
    param([System.Windows.Forms.Panel]$parentPanel)
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Rename Volume"
    $titleLabel.Location = New-Object System.Drawing.Point(0, 10)
    $titleLabel.Size = New-Object System.Drawing.Size(760, 30)
    $titleLabel.ForeColor = [System.Drawing.Color]::Lime
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $parentPanel.Controls.Add($titleLabel)
}

function New-RenameVolumeGroupBox {
    param([System.Windows.Forms.Panel]$parentPanel, [System.Windows.Forms.ListBox]$driveListBox)
    $groupBox = New-Object System.Windows.Forms.GroupBox
    $groupBox.Text = "Volume Rename Configuration"
    $groupBox.Location = New-Object System.Drawing.Point(180, 60)
    $groupBox.Size = New-Object System.Drawing.Size(400, 150)
    $groupBox.ForeColor = [System.Drawing.Color]::Lime
    $groupBox.BackColor = [System.Drawing.Color]::Transparent
    $groupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $parentPanel.Controls.Add($groupBox)

    # Drive letter label
    $driveLetterLabel = New-Object System.Windows.Forms.Label
    $driveLetterLabel.Text = "Drive Letter:"
    $driveLetterLabel.Location = New-Object System.Drawing.Point(30, 30)
    $driveLetterLabel.Size = New-Object System.Drawing.Size(100, 20)
    $driveLetterLabel.ForeColor = [System.Drawing.Color]::White
    $driveLetterLabel.BackColor = [System.Drawing.Color]::Transparent
    $driveLetterLabel.Font = New-Object System.Drawing.Font("Arial", 10)
    $groupBox.Controls.Add($driveLetterLabel)

    # Drive letter textbox - use script scope
    $script:renameDriveLetterTextBox = New-Object System.Windows.Forms.TextBox
    $script:renameDriveLetterTextBox.Location = New-Object System.Drawing.Point(130, 30)
    $script:renameDriveLetterTextBox.Size = New-Object System.Drawing.Size(50, 20)
    $script:renameDriveLetterTextBox.BackColor = [System.Drawing.Color]::Black
    $script:renameDriveLetterTextBox.ForeColor = [System.Drawing.Color]::Lime
    $script:renameDriveLetterTextBox.Font = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
    $script:renameDriveLetterTextBox.MaxLength = 1
    $script:renameDriveLetterTextBox.ReadOnly = $true
    $script:renameDriveLetterTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
    $groupBox.Controls.Add($script:renameDriveLetterTextBox)

    # New label label
    $newLabelLabel = New-Object System.Windows.Forms.Label
    $newLabelLabel.Text = "New Label:"
    $newLabelLabel.Location = New-Object System.Drawing.Point(30, 60)
    $newLabelLabel.Size = New-Object System.Drawing.Size(100, 20)
    $newLabelLabel.ForeColor = [System.Drawing.Color]::White
    $newLabelLabel.BackColor = [System.Drawing.Color]::Transparent
    $newLabelLabel.Font = New-Object System.Drawing.Font("Arial", 10)
    $groupBox.Controls.Add($newLabelLabel)

    # New label textbox - use script scope
    $script:renameNewLabelTextBox = New-Object System.Windows.Forms.TextBox
    $script:renameNewLabelTextBox.Location = New-Object System.Drawing.Point(130, 60)
    $script:renameNewLabelTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $script:renameNewLabelTextBox.BackColor = [System.Drawing.Color]::Black
    $script:renameNewLabelTextBox.ForeColor = [System.Drawing.Color]::Lime
    $script:renameNewLabelTextBox.Font = New-Object System.Drawing.Font("Consolas", 11)
    $groupBox.Controls.Add($script:renameNewLabelTextBox)

    return @{
        GroupBox           = $groupBox
        DriveLetterTextBox = $script:renameDriveLetterTextBox
        NewLabelTextBox    = $script:renameNewLabelTextBox
    }
}

function New-RenameActionButton {
    param([System.Windows.Forms.GroupBox]$groupBox, [System.Windows.Forms.ListBox]$driveListBox)
    $renameButton = New-DynamicButton -text "Rename" -x 100 -y 100 -width 200 -height 40 -clickAction {
        if ($script:renameDriveLetterTextBox -and $script:renameNewLabelTextBox) {
            $dl = $script:renameDriveLetterTextBox.Text.Trim().ToUpper()
            $nl = $script:renameNewLabelTextBox.Text.Trim()
            if ($dl -and $nl) {
                try {
                    Set-Volume -DriveLetter $dl -NewFileSystemLabel $nl -ErrorAction Stop
                    Add-Status "Renamed drive $dl to $nl successfully."
                }
                catch {
                    Add-Status "Error renaming drive: $_"
                }
            }
            else {
                Add-Status "Please enter both drive letter and new label."
            }
        }
    }
    $groupBox.Controls.Add($renameButton)
}

# [4.2] SHRINK VOLUME FUNCTIONS
function New-ShrinkVolumeTitle {
    param([System.Windows.Forms.Panel]$contentPanel)

    # Title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Shrink Volume and Create New Partition"
    $titleLabel.Location = New-Object System.Drawing.Point(0, 10)
    $titleLabel.Size = New-Object System.Drawing.Size(760, 30)
    $titleLabel.ForeColor = [System.Drawing.Color]::Lime
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $contentPanel.Controls.Add($titleLabel)
}

function New-ShrinkVolumeDriveSelector {
    param([System.Windows.Forms.Panel]$contentPanel)

    # Selected drive letter label
    $selectedDriveLabel = New-Object System.Windows.Forms.Label
    $selectedDriveLabel.Text = "Selected Drive Letter:"
    $selectedDriveLabel.Location = New-Object System.Drawing.Point(20, 50)
    $selectedDriveLabel.Size = New-Object System.Drawing.Size(150, 20)
    $selectedDriveLabel.ForeColor = [System.Drawing.Color]::White
    $selectedDriveLabel.BackColor = [System.Drawing.Color]::Transparent
    $selectedDriveLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $contentPanel.Controls.Add($selectedDriveLabel)

    # Selected drive letter textbox - use script scope
    $script:selectedDriveTextBox = New-Object System.Windows.Forms.TextBox
    $script:selectedDriveTextBox.Location = New-Object System.Drawing.Point(180, 50)
    $script:selectedDriveTextBox.Size = New-Object System.Drawing.Size(50, 25)
    $script:selectedDriveTextBox.BackColor = [System.Drawing.Color]::Black
    $script:selectedDriveTextBox.ForeColor = [System.Drawing.Color]::Lime
    $script:selectedDriveTextBox.Font = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
    $script:selectedDriveTextBox.MaxLength = 1
    $script:selectedDriveTextBox.ReadOnly = $true
    $script:selectedDriveTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
    $contentPanel.Controls.Add($script:selectedDriveTextBox)
}

function New-ShrinkVolumePartitionSizeOptions {
    param([System.Windows.Forms.Panel]$contentPanel)

    # Partition size options group box
    $partitionGroupBox = New-Object System.Windows.Forms.GroupBox
    $partitionGroupBox.Text = "Choose Partition Size"
    $partitionGroupBox.Location = New-Object System.Drawing.Point(20, 80)
    $partitionGroupBox.Size = New-Object System.Drawing.Size(720, 120)
    $partitionGroupBox.ForeColor = [System.Drawing.Color]::Lime
    $partitionGroupBox.BackColor = [System.Drawing.Color]::Transparent
    $partitionGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $contentPanel.Controls.Add($partitionGroupBox)

    # Create a panel inside the GroupBox to properly group radio buttons
    $radioPanel = New-Object System.Windows.Forms.Panel
    $radioPanel.Location = New-Object System.Drawing.Point(10, 20)
    $radioPanel.Size = New-Object System.Drawing.Size(700, 90)
    $radioPanel.BackColor = [System.Drawing.Color]::Transparent
    $partitionGroupBox.Controls.Add($radioPanel)

    # Declare radio buttons at script scope so they're accessible in the shrink button click event
    $script:radio80GB = New-Object System.Windows.Forms.RadioButton
    $script:radio80GB.Text = "100GB (recommended for 256GB drives)"
    $script:radio80GB.Location = New-Object System.Drawing.Point(10, 10)
    $script:radio80GB.Size = New-Object System.Drawing.Size(350, 20)
    $script:radio80GB.ForeColor = [System.Drawing.Color]::White
    $script:radio80GB.Font = New-Object System.Drawing.Font("Arial", 10)
    $script:radio80GB.Checked = $true
    $radioPanel.Controls.Add($script:radio80GB)

    # 200GB radio button
    $script:radio200GB = New-Object System.Windows.Forms.RadioButton
    $script:radio200GB.Text = "200GB (recommended for 500GB drives)"
    $script:radio200GB.Location = New-Object System.Drawing.Point(10, 35)
    $script:radio200GB.Size = New-Object System.Drawing.Size(350, 20)
    $script:radio200GB.ForeColor = [System.Drawing.Color]::White
    $script:radio200GB.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioPanel.Controls.Add($script:radio200GB)

    # 500GB radio button
    $script:radio500GB = New-Object System.Windows.Forms.RadioButton
    $script:radio500GB.Text = "500GB (recommended for 1TB+ drives)"
    $script:radio500GB.Location = New-Object System.Drawing.Point(10, 60)
    $script:radio500GB.Size = New-Object System.Drawing.Size(350, 20)
    $script:radio500GB.ForeColor = [System.Drawing.Color]::White
    $script:radio500GB.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioPanel.Controls.Add($script:radio500GB)

    # Custom size radio button
    $script:radioCustom = New-Object System.Windows.Forms.RadioButton
    $script:radioCustom.Text = "Custom size (MB):"
    $script:radioCustom.Location = New-Object System.Drawing.Point(370, 10)
    $script:radioCustom.Size = New-Object System.Drawing.Size(150, 20)
    $script:radioCustom.ForeColor = [System.Drawing.Color]::White
    $script:radioCustom.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioPanel.Controls.Add($script:radioCustom)

    # Custom size textbox
    $script:customSizeTextBox = New-Object System.Windows.Forms.TextBox
    $script:customSizeTextBox.Location = New-Object System.Drawing.Point(370, 35)
    $script:customSizeTextBox.Size = New-Object System.Drawing.Size(150, 25)
    $script:customSizeTextBox.BackColor = [System.Drawing.Color]::Black
    $script:customSizeTextBox.ForeColor = [System.Drawing.Color]::Lime
    $script:customSizeTextBox.Font = New-Object System.Drawing.Font("Consolas", 11)
    $script:customSizeTextBox.Text = "102400"  # Default to 100GB in MB
    $script:customSizeTextBox.Enabled = $false
    $radioPanel.Controls.Add($script:customSizeTextBox)

    # Add event handlers for radio buttons to enable/disable custom textbox
    $script:radioCustom.Add_CheckedChanged({
            if ($script:radioCustom.Checked) {
                $script:customSizeTextBox.Enabled = $true
                $script:customSizeTextBox.Focus()
            }
            else {
                $script:customSizeTextBox.Enabled = $false
            }
        })

    # Add event handlers for other radio buttons to disable custom textbox
    $script:radio80GB.Add_CheckedChanged({
            if ($script:radio80GB.Checked) {
                $script:customSizeTextBox.Enabled = $false
            }
        })

    $script:radio200GB.Add_CheckedChanged({
            if ($script:radio200GB.Checked) {
                $script:customSizeTextBox.Enabled = $false
            }
        })

    $script:radio500GB.Add_CheckedChanged({
            if ($script:radio500GB.Checked) {
                $script:customSizeTextBox.Enabled = $false
            }
        })
}

function New-ShrinkVolumeNewLabelInput {
    param([System.Windows.Forms.Panel]$contentPanel)

    # New partition label
    $newLabelLabel = New-Object System.Windows.Forms.Label
    $newLabelLabel.Text = "New Partition Label:"
    $newLabelLabel.Location = New-Object System.Drawing.Point(300, 50)
    $newLabelLabel.Size = New-Object System.Drawing.Size(150, 20)
    $newLabelLabel.ForeColor = [System.Drawing.Color]::White
    $newLabelLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $contentPanel.Controls.Add($newLabelLabel)

    # New partition label textbox
    $script:newLabelTextBox = New-Object System.Windows.Forms.TextBox
    $script:newLabelTextBox.Location = New-Object System.Drawing.Point(450, 50)
    $script:newLabelTextBox.Size = New-Object System.Drawing.Size(250, 25)
    $script:newLabelTextBox.BackColor = [System.Drawing.Color]::Black
    $script:newLabelTextBox.ForeColor = [System.Drawing.Color]::Lime
    $script:newLabelTextBox.Font = New-Object System.Drawing.Font("Consolas", 11)
    $script:newLabelTextBox.Text = "GAME"
    $contentPanel.Controls.Add($script:newLabelTextBox)
}

function Get-ShrinkVolumePartitionSize {
    # Determine partition size based on selected radio button
    $sizeMB = 0

    if ($script:radio80GB.Checked) {
        $sizeMB = 82020
    }
    elseif ($script:radio200GB.Checked) {
        $sizeMB = 204955
    }
    elseif ($script:radio500GB.Checked) {
        $sizeMB = 512000
    }
    elseif ($script:radioCustom.Checked) {
        # Validate custom size input
        $customSize = $script:customSizeTextBox.Text.Trim()
        if ($customSize -match '^\d+$') {
            try {
                $sizeMB = [int]$customSize
                if ($sizeMB -lt 1024) {
                    Add-Status "Error: Custom size must be at least 1024 MB (1 GB)."
                    return -1
                }
                if ($sizeMB -gt 2097152) {
                    # 2TB limit
                    Add-Status "Error: Custom size cannot exceed 2,097,152 MB (2 TB)."
                    return -1
                }
            }
            catch {
                Add-Status "Error processing custom size: $_"
                return -1
            }
        }
        else {
            Add-Status "Error: Custom size must be a valid number (digits only)."
            return -1
        }
    }
    else {
        Add-Status "Error: Please select a partition size option."
        return -1
    }

    return $sizeMB
}

function Test-ShrinkVolumeSpace {
    param([string]$driveLetter, [int]$sizeMB)

    # Validate drive exists and get info
    try {
        $driveInfo = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DeviceID -eq "$($driveLetter):" }
        if (-not $driveInfo) {
            Add-Status "Error: Drive $driveLetter does not exist."
            return $false
        }

        $freeSpaceMB = [math]::Floor($driveInfo.FreeSpace / 1MB)
        $totalSizeMB = [math]::Floor($driveInfo.Size / 1MB)

        # Get actual shrinkable space using PowerShell (more accurate)
        try {
            $partition = Get-Partition -DriveLetter $driveLetter -ErrorAction Stop
            $shrinkInfo = Get-PartitionSupportedSize -DriveLetter $driveLetter -ErrorAction Stop
            $maxShrinkBytes = $partition.Size - $shrinkInfo.SizeMin
            $maxShrinkMB = [math]::Floor($maxShrinkBytes / 1MB)

            if ($sizeMB -gt $maxShrinkMB) {
                Add-Status "Error: Requested size ($sizeMB MB) exceeds maximum shrinkable space ($maxShrinkMB MB)."
                Add-Status "Try running disk defragmentation first or choose a smaller size."
                return $false
            }
        }
        catch {
            # Fallback: Use 80% of free space as safe shrink limit
            $maxShrinkMB = [math]::Floor($freeSpaceMB * 0.8)

            if ($sizeMB -gt $maxShrinkMB) {
                Add-Status "Error: Requested size ($sizeMB MB) exceeds estimated safe shrink limit ($maxShrinkMB MB)."
                Add-Status "Try a smaller size or free up more space on the drive."
                return $false
            }
        }
    }
    catch {
        Add-Status "Error getting drive information: $_"
        return $false
    }

    return $true
}

function Invoke-ShrinkVolumeOperation {
    param([string]$driveLetter, [int]$sizeMB, [string]$newLabel)

    # Create a batch file that will run diskpart (using exact install.ps1 approach)
    $batchFilePath = "shrink_volume.bat"

    $batchContent = @"
@echo off
echo ============================================================ > shrink_status.txt
echo                  Shrinking Volume $driveLetter >> shrink_status.txt
echo ============================================================ >> shrink_status.txt
echo. >> shrink_status.txt

echo Creating diskpart script... >> shrink_status.txt
(
    echo select volume $driveLetter
    echo shrink desired=$sizeMB
    echo create partition primary
    echo format fs=ntfs quick
    echo assign
    echo list volume
) > diskpart_script.txt

echo Running diskpart... >> shrink_status.txt
echo. >> shrink_status.txt
echo Diskpart script contents: >> shrink_status.txt
type diskpart_script.txt >> shrink_status.txt
echo. >> shrink_status.txt

diskpart /s diskpart_script.txt > diskpart_output.txt
if %errorlevel% neq 0 (
    echo Error: Diskpart failed with exit code %errorlevel% >> shrink_status.txt
    echo This could be due to insufficient free space or the drive being in use. >> shrink_status.txt
    echo Try defragmenting the drive first or closing any applications using the drive. >> shrink_status.txt
    echo. >> shrink_status.txt
    echo Diskpart output: >> shrink_status.txt
    type diskpart_output.txt >> shrink_status.txt

    echo. >> shrink_status.txt
    echo Checking drive information: >> shrink_status.txt
    powershell -command "Get-WmiObject Win32_LogicalDisk -Filter \"DeviceID='$($driveLetter):'\" | Select-Object DeviceID, VolumeName, Size, FreeSpace | Format-List" >> shrink_status.txt

    del diskpart_output.txt
    del diskpart_script.txt
    exit /b %errorlevel%
)

echo Diskpart completed successfully. >> shrink_status.txt
echo. >> shrink_status.txt

echo Cleaning up temporary files... >> shrink_status.txt
del diskpart_output.txt
del diskpart_script.txt

echo. >> shrink_status.txt
echo Getting available drives after operation... >> shrink_status.txt
powershell -command "Get-WmiObject Win32_LogicalDisk | Select-Object @{Name='Name';Expression={`$_.DeviceID}}, @{Name='VolumeName';Expression={`$_.VolumeName}}, @{Name='Size (GB)';Expression={[math]::round(`$_.Size/1GB, 0)}}, @{Name='FreeSpace (GB)';Expression={[math]::round(`$_.FreeSpace/1GB, 0)}} | Format-Table -AutoSize | Out-String" >> shrink_status.txt

echo Operation completed successfully. >> shrink_status.txt
"@
    Set-Content -Path $batchFilePath -Value $batchContent -Force -Encoding ASCII

    Add-Status "Shrinking drive $driveLetter..."

    try {
        # Create a process to run batch file with admin privileges and hide cmd window
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "cmd.exe"
        $psi.Arguments = "/c `"$batchFilePath`""
        $psi.UseShellExecute = $true
        $psi.Verb = "runas"
        $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

        # Run process
        $batchProcess = [System.Diagnostics.Process]::Start($psi)
        $batchProcess.WaitForExit()

        # Read status file and display in status box
        if (Test-Path "shrink_status.txt") {
            $statusContent = Get-Content "shrink_status.txt" -Raw
            Remove-Item "shrink_status.txt" -Force -ErrorAction SilentlyContinue
        }

        # Check if operation was successful (using exact install.ps1 logic)
        if ($batchProcess.ExitCode -eq 0) {
            Add-Status "Operation completed successfully."
            Add-Status "Creating new partition..."

            # Refresh drive list
            Start-Sleep -Seconds 2

            # Find newly created drive (exact same logic as install.ps1)
            $newDriveFound = $false
            $newDriveLetter = ""

            # Wait a bit to ensure system has updated
            Start-Sleep -Seconds 2

            # Find newly created drive
            $currentDrives = Get-WmiObject Win32_LogicalDisk | Select-Object DeviceID, VolumeName
            foreach ($drive in $currentDrives) {
                if ($drive.DeviceID -ne "$($driveLetter):" -and
                    ($drive.VolumeName -eq "New Volume" -or $drive.VolumeName -eq "")) {
                    $newDriveFound = $true
                    $newDriveLetter = $drive.DeviceID.TrimEnd(":")
                    break
                }
            }

            # Rename the new drive if found
            if ($newDriveFound) {
                $actualNewLabel = if (-not [string]::IsNullOrEmpty($newLabel)) { $newLabel } else { "GAME" }

                # Rename using the most reliable method (Set-Volume)
                try {
                    Set-Volume -DriveLetter $newDriveLetter -NewFileSystemLabel $actualNewLabel -ErrorAction Stop
                    Add-Status "Successfully renamed drive $newDriveLetter to $actualNewLabel."
                }
                catch {
                    # Fallback to label command
                    try {
                        Start-Process -FilePath "cmd.exe" -ArgumentList "/c label $newDriveLetter`:$actualNewLabel" -WindowStyle Hidden -Wait
                        Add-Status "Successfully renamed drive $newDriveLetter to $actualNewLabel."
                    }
                    catch {
                        Add-Status "Failed to rename drive $newDriveLetter. Please rename manually to '$actualNewLabel'."
                    }
                }
            }
            else {
                Add-Status "Could not find the newly created drive. Please rename it manually."
            }

            # Update drive list
            $driveCount = Update-DriveList
            Add-Status "Drive list updated. Found $driveCount drives."
        }
        else {
            Add-Status "Operation completed with warnings. Check the event logs for details."
        }

        Remove-Item $batchFilePath -Force -ErrorAction SilentlyContinue
    }
    catch {
        Add-Status "Error: $_"
        Add-Status "Make sure you have administrator privileges."
        Remove-Item $batchFilePath -Force -ErrorAction SilentlyContinue
    }
}

function New-ShrinkVolumeActionButton {
    param([System.Windows.Forms.Panel]$contentPanel)

    # Shrink button
    $shrinkButton = New-DynamicButton -text "Shrink" -x 275 -y 210 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        $driveLetter = $script:selectedDriveTextBox.Text.Trim().ToUpper()
        $newLabel = $script:newLabelTextBox.Text.Trim()

        # Validate input
        if ($driveLetter -eq "") {
            Add-Status "Error: Please select a drive."
            return
        }

        if ($newLabel -eq "") {
            Add-Status "Error: Please enter a label for the new partition."
            return
        }

        # Get partition size
        $sizeMB = Get-ShrinkVolumePartitionSize
        if ($sizeMB -eq -1) {
            return  # Error already displayed
        }

        # Test if shrink operation is possible
        if (-not (Test-ShrinkVolumeSpace -driveLetter $driveLetter -sizeMB $sizeMB)) {
            return  # Error already displayed
        }

        # Perform shrink operation
        Invoke-ShrinkVolumeOperation -driveLetter $driveLetter -sizeMB $sizeMB -newLabel $newLabel
    }
    $contentPanel.Controls.Add($shrinkButton)
}

# [4.4] EXTEND VOLUME FUNCTIONS
function New-ExtendVolumeTitle {
    param([System.Windows.Forms.Panel]$parentPanel)

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Extend Volume by Merging"
    $titleLabel.Location = New-Object System.Drawing.Point(0, 10)
    $titleLabel.Size = New-Object System.Drawing.Size(760, 30)
    $titleLabel.ForeColor = [System.Drawing.Color]::Lime
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $parentPanel.Controls.Add($titleLabel)
}

function New-ExtendVolumeGroupBox {
    param([System.Windows.Forms.Panel]$parentPanel)

    # Create GroupBox for centered content
    $extendGroupBox = New-Object System.Windows.Forms.GroupBox
    $extendGroupBox.Text = "Volume Merge Configuration"  # ✅ Thêm Text để event handler có thể tìm thấy
    $extendGroupBox.Location = New-Object System.Drawing.Point(180, 60)
    $extendGroupBox.Size = New-Object System.Drawing.Size(400, 180)
    $extendGroupBox.ForeColor = [System.Drawing.Color]::Lime
    $extendGroupBox.BackColor = [System.Drawing.Color]::Transparent
    $extendGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $parentPanel.Controls.Add($extendGroupBox)

    # Source drive label
    $sourceDriveLabel = New-Object System.Windows.Forms.Label
    $sourceDriveLabel.Text = "Source Drive (to delete):"
    $sourceDriveLabel.Location = New-Object System.Drawing.Point(75, 35)
    $sourceDriveLabel.Size = New-Object System.Drawing.Size(180, 20)
    $sourceDriveLabel.ForeColor = [System.Drawing.Color]::White
    $sourceDriveLabel.BackColor = [System.Drawing.Color]::Transparent
    $sourceDriveLabel.Font = New-Object System.Drawing.Font("Arial", 10)
    $extendGroupBox.Controls.Add($sourceDriveLabel)

    # Source drive textbox - use script scope
    $script:extendSourceDriveTextBox = New-Object System.Windows.Forms.TextBox
    $script:extendSourceDriveTextBox.Location = New-Object System.Drawing.Point(260, 30)
    $script:extendSourceDriveTextBox.Size = New-Object System.Drawing.Size(60, 25)
    $script:extendSourceDriveTextBox.BackColor = [System.Drawing.Color]::Black
    $script:extendSourceDriveTextBox.ForeColor = [System.Drawing.Color]::Lime
    $script:extendSourceDriveTextBox.Font = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Bold)
    $script:extendSourceDriveTextBox.MaxLength = 1
    $script:extendSourceDriveTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
    # Add focus events for better user experience
    $script:extendSourceDriveTextBox.Add_GotFocus({ $this.SelectAll() })
    $script:extendSourceDriveTextBox.Add_TextChanged({ $this.Text = $this.Text.ToUpper() })
    $extendGroupBox.Controls.Add($script:extendSourceDriveTextBox)

    # Target drive label
    $targetDriveLabel = New-Object System.Windows.Forms.Label
    $targetDriveLabel.Text = "Target Drive (to expand):"
    $targetDriveLabel.Location = New-Object System.Drawing.Point(75, 65)
    $targetDriveLabel.Size = New-Object System.Drawing.Size(180, 20)
    $targetDriveLabel.ForeColor = [System.Drawing.Color]::White
    $targetDriveLabel.BackColor = [System.Drawing.Color]::Transparent
    $targetDriveLabel.Font = New-Object System.Drawing.Font("Arial", 10)
    $extendGroupBox.Controls.Add($targetDriveLabel)

    # Target drive textbox - use script scope
    $script:extendTargetDriveTextBox = New-Object System.Windows.Forms.TextBox
    $script:extendTargetDriveTextBox.Location = New-Object System.Drawing.Point(260, 60)
    $script:extendTargetDriveTextBox.Size = New-Object System.Drawing.Size(60, 25)
    $script:extendTargetDriveTextBox.BackColor = [System.Drawing.Color]::Black
    $script:extendTargetDriveTextBox.ForeColor = [System.Drawing.Color]::Lime
    $script:extendTargetDriveTextBox.Font = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Bold)
    $script:extendTargetDriveTextBox.MaxLength = 1
    $script:extendTargetDriveTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
    # Add focus events for better user experience
    $script:extendTargetDriveTextBox.Add_GotFocus({ $this.SelectAll() })
    $script:extendTargetDriveTextBox.Add_TextChanged({ $this.Text = $this.Text.ToUpper() })
    $extendGroupBox.Controls.Add($script:extendTargetDriveTextBox)

    # Warning label
    $warningLabel = New-Object System.Windows.Forms.Label
    $warningLabel.Text = "WARNING: This will DELETE drive and all data!"
    $warningLabel.Location = New-Object System.Drawing.Point(30, 100)
    $warningLabel.Size = New-Object System.Drawing.Size(340, 25)
    $warningLabel.ForeColor = [System.Drawing.Color]::Red
    $warningLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $warningLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $extendGroupBox.Controls.Add($warningLabel)

    # Add Enter key navigation
    $script:extendSourceDriveTextBox.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true
                $script:extendTargetDriveTextBox.Focus()
            }
        })

    return @{
        GroupBox           = $extendGroupBox
        SourceDriveTextBox = $script:extendSourceDriveTextBox
        TargetDriveTextBox = $script:extendTargetDriveTextBox
    }
}

function New-ExtendActionButton {
    param([hashtable]$extendControls)

    $groupBox = $extendControls.GroupBox

    # Extend button (inside GroupBox)
    $mergeButton = New-DynamicButton -text "Extend" -x 100 -y 130 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Get values directly from script scope variables
        $sourceDrive = ""
        $targetDrive = ""

        if ($script:extendSourceDriveTextBox -and $script:extendSourceDriveTextBox.Text) {
            $sourceDrive = $script:extendSourceDriveTextBox.Text.Trim().ToUpper()
        }
        if ($script:extendTargetDriveTextBox -and $script:extendTargetDriveTextBox.Text) {
            $targetDrive = $script:extendTargetDriveTextBox.Text.Trim().ToUpper()
        }

        Add-Status "Source Drive: '$sourceDrive' | Target Drive: '$targetDrive'"

        # Validate input
        if (-not (Test-ExtendVolumeInput -sourceDrive $sourceDrive -targetDrive $targetDrive)) {
            return
        }

        # Confirm operation
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "WARNING: This will DELETE drive $sourceDrive and all its data, then extend drive $targetDrive.`n`nAre you sure you want to continue?",
            "Confirm Merge",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::No) {
            Add-Status "Operation cancelled by user."
            return
        }

        # Perform merge operation using script scope textboxes
        Add-Status "Merging volumes: deleting drive $sourceDrive and extending drive $targetDrive..."
        Invoke-ExtendVolumeOperation -sourceDrive $sourceDrive -targetDrive $targetDrive -sourceDriveTextBox $script:extendSourceDriveTextBox -targetDriveTextBox $script:extendTargetDriveTextBox
    }
    $groupBox.Controls.Add($mergeButton)

    # Set up Enter key for target textbox to trigger merge
    $script:extendTargetDriveTextBox.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true
                $mergeButton.PerformClick()
            }
        })

    return $mergeButton
}

function Test-ExtendVolumeInput {
    param([string]$sourceDrive, [string]$targetDrive)

    # Basic validation
    if ([string]::IsNullOrEmpty($sourceDrive)) {
        Add-Status "Error: Please enter a source drive letter."
        return $false
    }
    if ([string]::IsNullOrEmpty($targetDrive)) {
        Add-Status "Error: Please enter a target drive letter."
        return $false
    }
    if (-not ($sourceDrive -match '^[A-Z]$') -or -not ($targetDrive -match '^[A-Z]$')) {
        Add-Status "Error: Drive letters must be single letters (A-Z)."
        return $false
    }
    if ($sourceDrive -eq $targetDrive) {
        Add-Status "Error: Source and target drives cannot be the same."
        return $false
    }

    # Check if drives exist
    $existingDrives = Get-CimInstance -ClassName Win32_LogicalDisk | Select-Object -ExpandProperty DeviceID | ForEach-Object { $_.Substring(0, 1) }
    if ($existingDrives -notcontains $sourceDrive) {
        Add-Status "Error: Source drive $sourceDrive does not exist."
        return $false
    }
    if ($existingDrives -notcontains $targetDrive) {
        Add-Status "Error: Target drive $targetDrive does not exist."
        return $false
    }

    # Check if drives are on same physical disk
    try {
        $sourcePartition = Get-Partition -DriveLetter $sourceDrive -ErrorAction Stop
        $targetPartition = Get-Partition -DriveLetter $targetDrive -ErrorAction Stop

        if ($sourcePartition.DiskNumber -ne $targetPartition.DiskNumber) {
            Add-Status "Error: Drives are not on the same physical disk. Operation aborted for safety."
            return $false
        }
    }
    catch {
        Add-Status "Warning: Could not verify disk compatibility. Proceeding anyway..."
    }

    return $true
}

function Invoke-ExtendVolumeOperation {
    param([string]$sourceDrive, [string]$targetDrive, $sourceDriveTextBox, $targetDriveTextBox)

    Add-Status "Starting volume merge operation..."

    # Kiểm tra xem hai ổ đĩa có nằm trên cùng một đĩa vật lý không
    try {
        # Phương pháp 1: Sử dụng Get-Partition
        try {
            $sourcePartition = Get-Partition -DriveLetter $sourceDrive -ErrorAction Stop
            $targetPartition = Get-Partition -DriveLetter $targetDrive -ErrorAction Stop

            $sourceDiskNumber = $sourcePartition.DiskNumber
            $targetDiskNumber = $targetPartition.DiskNumber

            if ($sourceDiskNumber -ne $targetDiskNumber) {
                Add-Status "Error: Drives are not on the same physical disk. Operation aborted for safety."
                return
            }
        }
        catch {
            Add-Status "Warning: Could not verify disk compatibility. Proceeding anyway..."
        }
    }
    catch {
        Add-Status "Warning: Could not verify if drives are on the same physical disk. Proceeding anyway..."
    }

    # Create a batch file that will run the merge operation
    $batchFilePath = "merge_volumes.bat"
    $batchContent = @"
@echo off
setlocal enabledelayedexpansion

echo ============================================================ > merge_log.txt
echo                  Merging Volumes >> merge_log.txt
echo ============================================================ >> merge_log.txt
echo. >> merge_log.txt

echo Deleting source drive $sourceDrive... >> merge_log.txt
powershell -WindowStyle Hidden -command "& { try { Remove-Partition -DriveLetter $sourceDrive -Confirm:`$false -ErrorAction Stop; Write-Output 'Successfully deleted source drive $sourceDrive.' } catch { Write-Error `$_.Exception.Message; exit 1 } }" > delete_output.txt 2>&1

type delete_output.txt >> merge_log.txt

if errorlevel 1 (
    echo PowerShell delete failed, trying diskpart... >> merge_log.txt
    (
        echo select volume $sourceDrive
        echo delete volume override
    ) > diskpart_delete.txt

    start /b /wait "" cmd /c "diskpart /s diskpart_delete.txt > diskpart_delete_output.txt 2>&1"
    type diskpart_delete_output.txt >> merge_log.txt

    if errorlevel 1 (
        echo ERROR: Failed to delete source drive $sourceDrive. >> merge_log.txt
        del diskpart_delete.txt
        del diskpart_delete_output.txt
        del delete_output.txt
        exit /b 1
    )
    del diskpart_delete.txt
    del diskpart_delete_output.txt
)
del delete_output.txt

echo Waiting for system to update... >> merge_log.txt
timeout /t 2 /nobreak > nul

echo Extending target drive $targetDrive... >> merge_log.txt
powershell -WindowStyle Hidden -command "& { try { `$size = (Get-PartitionSupportedSize -DriveLetter $targetDrive).SizeMax; Resize-Partition -DriveLetter $targetDrive -Size `$size -ErrorAction Stop; Write-Output 'Successfully extended partition using PowerShell.' } catch { Write-Error `$_.Exception.Message; exit 1 } }" > extend_output.txt 2>&1

type extend_output.txt >> merge_log.txt

if errorlevel 1 (
    echo PowerShell extend failed, trying diskpart... >> merge_log.txt
    (
        echo rescan
        echo select volume $targetDrive
        echo extend
    ) > diskpart_extend.txt

    start /b /wait "" cmd /c "diskpart /s diskpart_extend.txt > diskpart_extend_output.txt 2>&1"
    type diskpart_extend_output.txt >> merge_log.txt

    if errorlevel 1 (
        echo ERROR: Failed to extend target drive $targetDrive. >> merge_log.txt
        del diskpart_extend.txt
        del diskpart_extend_output.txt
        del extend_output.txt
        exit /b 1
    )
    del diskpart_extend.txt
    del diskpart_extend_output.txt
) else (
    echo Successfully extended partition using PowerShell. >> merge_log.txt
)
del extend_output.txt

echo. >> merge_log.txt
echo Merge completed successfully! >> merge_log.txt
echo Operation completed. >> merge_log.txt
exit /b 0
"@

    Set-Content -Path $batchFilePath -Value $batchContent -Force -Encoding ASCII

    Add-Status "Merging volumes: deleting drive $sourceDrive and extending drive $targetDrive..."
    Add-Status "Processing... Please wait while the operation completes."

    try {
        # Create a process to run batch file with admin privileges and hide cmd window
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "cmd.exe"
        $psi.Arguments = "/c `"$batchFilePath`""
        $psi.UseShellExecute = $true
        $psi.Verb = "runas"
        $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

        # Run process
        $batchProcess = [System.Diagnostics.Process]::Start($psi)

        # Show progress while waiting
        $progressCounter = 0
        $progressChars = @('|', '/', '-', '\')
        $progressSteps = @(
            "Deleting source drive...",
            "Waiting for system update...",
            "Extending target drive...",
            "Finalizing operation..."
        )
        $currentStep = 0
        $stepDuration = 0
        $maxStepDuration = 8

        while (!$batchProcess.HasExited) {
            $progressChar = $progressChars[$progressCounter % 4]
            $stepDuration++
            if ($stepDuration -ge $maxStepDuration) {
                $currentStep = ($currentStep + 1) % $progressSteps.Count
                $stepDuration = 0
            }

            $currentMessage = $progressSteps[$currentStep]
            $progressCounter++
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 250
        }

        # Check if operation was successful
        if ($batchProcess.ExitCode -eq 0) {
            Add-Status "Operation completed successfully."
            Add-Status "Merged volumes: deleted drive $sourceDrive and extended drive $targetDrive."

            # Update drive list
            $driveCount = Update-DriveList
            Add-Status "Drive list updated. Found $driveCount drives."

            # Clear textboxes
            $sourceDriveTextBox.Text = ""
            $targetDriveTextBox.Text = ""
        }
        else {
            Add-Status "Operation completed with warnings or errors."
            Add-Status "Exit code: $($batchProcess.ExitCode)"
        }

        # Clean up files
        Remove-Item $batchFilePath -Force -ErrorAction SilentlyContinue
        Remove-Item "merge_log.txt" -Force -ErrorAction SilentlyContinue
    }
    catch {
        Add-Status "Error: $_"
        Add-Status "Make sure you have administrator privileges."
        Remove-Item $batchFilePath -Force -ErrorAction SilentlyContinue
    }
}

# MAIN VOLUME MANAGEMENT DIALOG
function Invoke-VolumeManagementDialog {
    Hide-MainMenu
    # Create volume management form
    $volumeForm = New-Object System.Windows.Forms.Form
    $volumeForm.Text = "Volume Management"
    $volumeForm.Size = New-Object System.Drawing.Size(795, 660) # Increase the size of the form
    $volumeForm.StartPosition = "CenterScreen"
    $volumeForm.BackColor = [System.Drawing.Color]::Black
    $volumeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $volumeForm.MaximizeBox = $false
    $volumeForm.MinimizeBox = $false

    # Add gradient background identical to main menu
    $volumeForm.Add_Paint({
        $graphics = $_.Graphics
        $rect = New-Object System.Drawing.Rectangle(0, 0, $volumeForm.Width, $volumeForm.Height)
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $rect,
            [System.Drawing.Color]::FromArgb(0, 0, 0), # Black at top
            [System.Drawing.Color]::FromArgb(0, 30, 0), # Dark green at bottom
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
        )
        $graphics.FillRectangle($brush, $rect)
        $brush.Dispose()
    })


    # Title label with animation
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "VOLUME MANAGEMENT"
    $titleLabel.Location = New-Object System.Drawing.Point(0, 10) # Move the title label down
    $titleLabel.Size = New-Object System.Drawing.Size(795, 40) # Increase the size of the title label
    $titleLabel.ForeColor = [System.Drawing.Color]::Lime
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $titleLabel.Padding = New-Object System.Windows.Forms.Padding(5)

    # Add animation to the title
    $titleTimer = New-Object System.Windows.Forms.Timer
    $titleTimer.Interval = 500
    $titleTimer.Add_Tick({
            if ($titleLabel.ForeColor -eq [System.Drawing.Color]::Lime) {
                $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 220, 0)
            }
            else {
                $titleLabel.ForeColor = [System.Drawing.Color]::Lime
            }
        })
    $titleTimer.Start()

    $volumeForm.Controls.Add($titleLabel)

    # Drive list box with enhanced styling
    $driveListBox = New-Object System.Windows.Forms.ListBox
    $driveListBox.Location = New-Object System.Drawing.Point(10, 50) # Move the drive list box down
    $driveListBox.Size = New-Object System.Drawing.Size(760, 100) # Increase the size of the drive list box
    $driveListBox.BackColor = [System.Drawing.Color]::Black
    $driveListBox.ForeColor = [System.Drawing.Color]::Lime
    $driveListBox.Font = New-Object System.Drawing.Font("Consolas", 10)
    $driveListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $volumeForm.Controls.Add($driveListBox)

    # Content Panel for function buttons
    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Location = New-Object System.Drawing.Point(10, 200)
    $contentPanel.Size = New-Object System.Drawing.Size(760, 260)
    $contentPanel.BackColor = [System.Drawing.Color]::Transparent
    $contentPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    # Add content panel to form
    $volumeForm.Controls.Add($contentPanel)

    # Status text box with enhanced styling
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(10, 470) # Move the status text box down
    $statusTextBox.Size = New-Object System.Drawing.Size(760, 140) # Increase the size of the status text box
    $statusTextBox.BackColor = [System.Drawing.Color]::Black
    $statusTextBox.ForeColor = [System.Drawing.Color]::Lime
    $statusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $statusTextBox.ReadOnly = $true
    $statusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $statusTextBox.Text = "Status messages will appear here..."
    $volumeForm.Controls.Add($statusTextBox)

    # Function to add status message with timestamp
    function Add-Status {
        param([string]$message)

        # Clear placeholder text on first message
        if ($statusTextBox.Text -eq "Status messages will appear here...") {
            $statusTextBox.Clear()
        }

        # Add timestamp to message
        $timestamp = Get-Date -Format "HH:mm:ss"
        $statusTextBox.AppendText("[$timestamp] $message`r`n")
        $statusTextBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    $driveCount = Update-DriveList

    # Add a common event handler for driveListBox to update all input fields in all buttons
    $driveListBox.Add_SelectedIndexChanged({
        if ($driveListBox.SelectedItem) {
            $selectedDrive = $driveListBox.SelectedItem.ToString()
            $driveLetter = $selectedDrive.Substring(0, 1)
            # Update for Change Letter button
            if ($contentPanel.Controls.Count -gt 0 -and $contentPanel.Controls[0].Text -eq "Change Drive Letter") {
                # Find the GroupBox in the change letter panel
                $changeGroupBox = $contentPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.GroupBox] }
                if ($changeGroupBox) {
                    # Find the old drive letter textbox (first textbox)
                    $oldLetterTextBox = $changeGroupBox.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] } | Select-Object -First 1
                    if ($oldLetterTextBox) {
                        $oldLetterTextBox.Text = $driveLetter
                    }
                }
            }
            # Update for Shrink Volume button
            if ($contentPanel.Controls.Count -gt 0 -and $contentPanel.Controls[0].Text -eq "Shrink Volume and Create New Partition") {
                # Use script scope variable for shrink volume
                if ($script:selectedDriveTextBox) {
                    $script:selectedDriveTextBox.Text = $driveLetter
                }
            }
            # Update for Extend Volume button
            if ($contentPanel.Controls.Count -gt 0 -and $contentPanel.Controls[0].Text -eq "Extend Volume by Merging") {
                # Find the textboxes in the extend volume panel
                $extendGroupBox = $contentPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.GroupBox] }
                if ($extendGroupBox) {
                    $textBoxes = $extendGroupBox.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] }
                    if ($textBoxes.Count -ge 2) {
                        $sourceTextBox = $textBoxes[0]
                        $targetTextBox = $textBoxes[1]

                        # If source drive is empty, fill it
                        if ($sourceTextBox.Text -eq "") {
                            $sourceTextBox.Text = $driveLetter
                        }
                        # Otherwise, if target drive is empty and different from source, fill it
                        elseif ($targetTextBox.Text -eq "" -and $driveLetter -ne $sourceTextBox.Text) {
                            $targetTextBox.Text = $driveLetter
                        }
                    }
                }
            }
            # Update for Rename Volume button
            if ($contentPanel.Controls.Count -gt 0 -and $contentPanel.Controls[0].Text -eq "Rename Volume") {
                # Find the GroupBox in the rename panel
                $renameGroupBox = $contentPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.GroupBox] }
                if ($renameGroupBox) {
                    # Find the drive letter textbox (first textbox)
                    $driveLetterTextBox = $renameGroupBox.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] } | Select-Object -First 1
                    if ($driveLetterTextBox) {
                        $driveLetterTextBox.Text = $driveLetter
                    }
                }
            }
        }
    })

    # [4.1] Change Drive Letter button
    $btnChangeDriveLetter = New-DynamicButton -text "Change Letter" -x 10 -y 150 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Clear the content panel
        $contentPanel.Controls.Clear()

        # Title label
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "Change Drive Letter"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 10)
        $titleLabel.Size = New-Object System.Drawing.Size(760, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $contentPanel.Controls.Add($titleLabel)

        # Create GroupBox for centered content
        $changeGroupBox = New-Object System.Windows.Forms.GroupBox
        $changeGroupBox.Text = "Drive Letter Configuration"
        $changeGroupBox.Location = New-Object System.Drawing.Point(180, 60)
        $changeGroupBox.Size = New-Object System.Drawing.Size(400, 150)
        $changeGroupBox.ForeColor = [System.Drawing.Color]::Lime
        $changeGroupBox.BackColor = [System.Drawing.Color]::Transparent
        $changeGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

        # Add GroupBox to content panel
        $contentPanel.Controls.Add($changeGroupBox)

        # Old drive letter label
        $oldLetterLabel = New-Object System.Windows.Forms.Label
        $oldLetterLabel.Text = "Select Drive Letter to Change:"
        $oldLetterLabel.Location = New-Object System.Drawing.Point(20, 30)
        $oldLetterLabel.Size = New-Object System.Drawing.Size(200, 20)
        $oldLetterLabel.ForeColor = [System.Drawing.Color]::White
        $oldLetterLabel.BackColor = [System.Drawing.Color]::Transparent
        $oldLetterLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $changeGroupBox.Controls.Add($oldLetterLabel)

        # Old drive letter textbox
        $script:oldLetterTextBox = New-Object System.Windows.Forms.TextBox
        $script:oldLetterTextBox.Location = New-Object System.Drawing.Point(230, 30)
        $script:oldLetterTextBox.Size = New-Object System.Drawing.Size(50, 20)
        $script:oldLetterTextBox.BackColor = [System.Drawing.Color]::Black
        $script:oldLetterTextBox.ForeColor = [System.Drawing.Color]::Lime
        $script:oldLetterTextBox.Font = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
        $script:oldLetterTextBox.MaxLength = 1
        $script:oldLetterTextBox.ReadOnly = $true
        $script:oldLetterTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
        $changeGroupBox.Controls.Add($script:oldLetterTextBox)

        # New drive letter label
        $newLetterLabel = New-Object System.Windows.Forms.Label
        $newLetterLabel.Text = "New Drive Letter:"
        $newLetterLabel.Location = New-Object System.Drawing.Point(20, 60)
        $newLetterLabel.Size = New-Object System.Drawing.Size(200, 20)
        $newLetterLabel.ForeColor = [System.Drawing.Color]::White
        $newLetterLabel.BackColor = [System.Drawing.Color]::Transparent
        $newLetterLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $changeGroupBox.Controls.Add($newLetterLabel)

        # New drive letter textbox
        $script:newLetterTextBox = New-Object System.Windows.Forms.TextBox
        $script:newLetterTextBox.Location = New-Object System.Drawing.Point(230, 60)
        $script:newLetterTextBox.Size = New-Object System.Drawing.Size(50, 20)
        $script:newLetterTextBox.BackColor = [System.Drawing.Color]::Black
        $script:newLetterTextBox.ForeColor = [System.Drawing.Color]::Lime
        $script:newLetterTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $script:newLetterTextBox.MaxLength = 1
        $script:newLetterTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
        $changeGroupBox.Controls.Add($script:newLetterTextBox)

        # Set initial value if a drive is already selected
        if ($driveListBox.SelectedItem) {
            $selectedDrive = $driveListBox.SelectedItem.ToString()
            $driveLetter = $selectedDrive.Substring(0, 1)
            $script:oldLetterTextBox.Text = $driveLetter
        }

        # Change button (inside GroupBox)
        $changeButton = New-DynamicButton -text "Change" -x 100 -y 100 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            $oldLetter = if ($script:oldLetterTextBox) { $script:oldLetterTextBox.Text.Trim().ToUpper() } else { "" }
            $newLetter = if ($script:newLetterTextBox) { $script:newLetterTextBox.Text.Trim().ToUpper() } else { "" }

            # Validate input
            if ($oldLetter -eq "") {
                Add-Status "Error: Please select a drive letter to change."
                return
            }

            if ($newLetter -eq "") {
                Add-Status "Error: Please enter a new drive letter."
                return
            }

            if ($oldLetter -eq $newLetter) {
                Add-Status "Error: New drive letter must be different from the current one."
                return
            }

            # Check if new letter is already in use
            $existingDrives = Get-WmiObject Win32_LogicalDisk | Select-Object -ExpandProperty DeviceID
            if ($existingDrives -contains "$($newLetter):") {
                Add-Status "Error: Drive letter $newLetter is already in use."
                return
            }

            # Create diskpart script
            $tempFile = [System.IO.Path]::GetTempFileName()
            $diskpartScript = @"
select volume $oldLetter
assign letter=$newLetter
"@
            Set-Content -Path $tempFile -Value $diskpartScript

            Add-Status "Changing drive $oldLetter to $newLetter..."

            try {
                # Run diskpart with elevated privileges
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = "diskpart.exe"
                $psi.Arguments = "/s `"$tempFile`""
                $psi.UseShellExecute = $true
                $psi.Verb = "runas"
                $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

                $process = [System.Diagnostics.Process]::Start($psi)
                $process.WaitForExit()

                # Check if successful
                if ($process.ExitCode -eq 0) {
                    Add-Status "Successfully changed drive letter from $oldLetter to $newLetter."

                    # Update drive list
                    $driveCount = Update-DriveList
                    Add-Status "Drive list updated. Found $driveCount drives."

                    # Clear textboxes
                    $script:oldLetterTextBox.Text = ""
                    $script:newLetterTextBox.Text = ""
                }
                else {
                    Add-Status "Error changing drive letter. Exit code: $($process.ExitCode)"
                }
            }
            catch {
                Add-Status "Error: $_"
            }
            finally {
                # Clean up temp file
                if (Test-Path $tempFile) {
                    Remove-Item $tempFile -Force
                }
            }
        }
        $changeGroupBox.Controls.Add($changeButton)

        # Set initial value if a drive is already selected
        if ($driveListBox.SelectedItem) {
            $selectedDrive = $driveListBox.SelectedItem.ToString()
            $driveLetter = $selectedDrive.Substring(0, 1)
            $script:oldLetterTextBox.Text = $driveLetter
        }
        Add-Status "Ready to change letter. Select a drive, enter a new letter, then click Change."
    }
    $volumeForm.Controls.Add($btnChangeDriveLetter)

    # [4.2] Shrink Volume button
    $btnShrinkVolume = New-DynamicButton -text "Shrink Volume" -x 170 -y 150 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Clear the content panel
        $contentPanel.Controls.Clear()

        # Create title using function
        New-ShrinkVolumeTitle -contentPanel $contentPanel

        # Create drive selector using function
        New-ShrinkVolumeDriveSelector -contentPanel $contentPanel

        # Create partition size options using function
        New-ShrinkVolumePartitionSizeOptions -contentPanel $contentPanel

        # Create new label input using function
        New-ShrinkVolumeNewLabelInput -contentPanel $contentPanel

        # Create shrink action button using function
        New-ShrinkVolumeActionButton -contentPanel $contentPanel

        # Update drive letter from selected drive IMMEDIATELY after controls are added
        if ($driveListBox.SelectedItem) {
            $selectedDrive = $driveListBox.SelectedItem.ToString()
            $driveLetter = $selectedDrive.Substring(0, 1)
            $script:selectedDriveTextBox.Text = $driveLetter
        }

        Add-Status "Ready to shrink volume. Select a drive, choose partition size, then click Shrink."
    }
    $volumeForm.Controls.Add($btnShrinkVolume)

    # [4.3] Rename Volume button
    $btnRenameVolume = New-DynamicButton -text "Rename Volume" -x 330 -y 150 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Clear the content panel
        $contentPanel.Controls.Clear()

        # Create title
        New-RenameVolumeTitle -parentPanel $contentPanel

        # Create GroupBox with all controls inside
        $renameControls = New-RenameVolumeGroupBox -parentPanel $contentPanel -driveListBox $driveListBox

        # Create rename button inside GroupBox
        New-RenameActionButton -groupBox $renameControls.GroupBox -driveListBox $driveListBox

        # ✅ Ensure drive letter is updated immediately when button is clicked
        if ($driveListBox.SelectedItem -and $renameControls.DriveLetterTextBox) {
            $selectedDrive = $driveListBox.SelectedItem.ToString()
            $driveLetter = $selectedDrive.Substring(0, 1)
            $renameControls.DriveLetterTextBox.Text = $driveLetter
        }

        Add-Status "Ready to rename volume. Select a drive, enter a new label, then click Rename Volume."
    }
    $volumeForm.Controls.Add($btnRenameVolume)

    # [4.4] Extend Volume button
    $btnExtendVolume = New-DynamicButton -text "Extend Volume" -x 490 -y 150 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Clear the content panel
        $contentPanel.Controls.Clear()

        # Create title
        New-ExtendVolumeTitle -parentPanel $contentPanel

        # Create GroupBox with all controls inside
        $extendControls = New-ExtendVolumeGroupBox -parentPanel $contentPanel

        # Create merge button inside GroupBox
        New-ExtendActionButton -extendControls $extendControls

        # ✅ Ensure drives are updated immediately when button is clicked
        if ($driveListBox.SelectedItem -and $script:extendSourceDriveTextBox) {
            $selectedDrive = $driveListBox.SelectedItem.ToString()
            $driveLetter = $selectedDrive.Substring(0, 1)
            $script:extendSourceDriveTextBox.Text = $driveLetter
        }

        Add-Status "Ready to extend volume. Select source and target drives, then click Extend."
    }
    $volumeForm.Controls.Add($btnExtendVolume)

    # [4.0] Return to Main Menu button
    $btnReturn = New-DynamicButton -text "Return" -x 650 -y 150 -width 120 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        $volumeForm.Close()
    }
    $volumeForm.Controls.Add($btnReturn)

    # When the form is closed, show the main menu again
    $volumeForm.Add_FormClosed({
            Show-MainMenu
        })

    # Set the cancel button (Escape key)
    $volumeForm.CancelButton = $btnReturn

    # Show the form
    $volumeForm.ShowDialog()
}

# [7] Rename Device Functions
function Rename-DeviceWithBatch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$newName,

        [scriptblock]$statusCallback,

        [bool]$showUI = $false
    )

    # Function to add status
    function Add-Status {
        param([string]$message)
        if ($statusCallback) {
            & $statusCallback $message
        }
    }

    try {
        # Validate computer name
        if ([string]::IsNullOrWhiteSpace($newName)) {
            $errorMsg = "Error: New computer name cannot be empty."
            Add-Status $errorMsg
            if ($showUI) {
                [System.Windows.Forms.MessageBox]::Show($errorMsg, "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            return $false
        }

        # Remove spaces and convert to uppercase
        $newName = $newName.Trim().ToUpper()

        # Validate computer name format
        if ($newName.Length -gt 15) {
            $errorMsg = "Error: Computer name cannot exceed 15 characters."
            Add-Status $errorMsg
            if ($showUI) {
                [System.Windows.Forms.MessageBox]::Show($errorMsg, "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            return $false
        }

        if ($newName -match '[^A-Z0-9-]') {
            $errorMsg = "Error: Computer name can only contain letters, numbers, and hyphens."
            Add-Status $errorMsg
            if ($showUI) {
                [System.Windows.Forms.MessageBox]::Show($errorMsg, "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            return $false
        }

        # Check if name is the same as current
        $currentName = $env:COMPUTERNAME
        if ($newName -eq $currentName) {
            $errorMsg = "Error: New name is the same as current name ($currentName)."
            Add-Status $errorMsg
            if ($showUI) {
                [System.Windows.Forms.MessageBox]::Show($errorMsg, "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            return $false
        }

        Add-Status "Validating new computer name: $newName"
        Add-Status "Current computer name: $currentName"

        # Confirm with user if showUI is enabled
        if ($showUI) {
            $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                "Are you sure you want to rename this computer from '$currentName' to '$newName'?`n`nThis will require a restart to take effect.",
                "Confirm Computer Rename",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
                Add-Status "Operation cancelled by user."
                return $false
            }
        }

        # Create batch file for renaming computer
        $batchFilePath = [System.IO.Path]::GetTempFileName() + ".bat"
        $batchContent = @"
@echo off
echo ============================================================ > rename_log.txt
echo              Computer Rename Operation >> rename_log.txt
echo ============================================================ >> rename_log.txt
echo. >> rename_log.txt
echo Current name: $currentName >> rename_log.txt
echo New name: $newName >> rename_log.txt
echo. >> rename_log.txt

echo Renaming computer to $newName... >> rename_log.txt
powershell -WindowStyle Hidden -Command "& { try { Rename-Computer -NewName '$newName' -Force -ErrorAction Stop; Write-Output 'Computer renamed successfully.' } catch { Write-Error `$_.Exception.Message; exit 1 } }" > rename_output.txt 2>&1

type rename_output.txt >> rename_log.txt

if errorlevel 1 (
    echo PowerShell rename failed, trying wmic... >> rename_log.txt
    wmic computersystem where name="%COMPUTERNAME%" call rename name="$newName" > wmic_output.txt 2>&1
    type wmic_output.txt >> rename_log.txt

    if errorlevel 1 (
        echo ERROR: Failed to rename computer using both PowerShell and WMIC. >> rename_log.txt
        del wmic_output.txt
        del rename_output.txt
        exit /b 1
    )
    del wmic_output.txt
) else (
    echo Successfully renamed computer using PowerShell. >> rename_log.txt
)
del rename_output.txt

echo. >> rename_log.txt
echo Computer rename completed successfully! >> rename_log.txt
echo A restart is required for the changes to take effect. >> rename_log.txt
exit /b 0
"@

        Set-Content -Path $batchFilePath -Value $batchContent -Force -Encoding ASCII

        Add-Status "Renaming computer from '$currentName' to '$newName'..."
        Add-Status "Processing... Please wait while the operation completes."

        # Create a process to run batch file with admin privileges
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "cmd.exe"
        $psi.Arguments = "/c `"$batchFilePath`""
        $psi.UseShellExecute = $true
        $psi.Verb = "runas"
        $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

        # Run process
        $batchProcess = [System.Diagnostics.Process]::Start($psi)
        $batchProcess.WaitForExit()

        # Check if operation was successful
        if ($batchProcess.ExitCode -eq 0) {
            Add-Status "Computer successfully renamed to '$newName'."
            Add-Status "A restart is required for the changes to take effect."

            if ($showUI) {
                $restartResult = [System.Windows.Forms.MessageBox]::Show(
                    "Computer has been successfully renamed to '$newName'.`n`nA restart is required for the changes to take effect.`n`nWould you like to restart now?",
                    "Rename Successful",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )

                if ($restartResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Add-Status "Restarting computer..."
                    Start-Process -FilePath "shutdown.exe" -ArgumentList "/r /t 5" -NoNewWindow
                }
            }

            $success = $true
        }
        else {
            Add-Status "Operation completed with warnings or errors."
            Add-Status "Exit code: $($batchProcess.ExitCode)"

            if ($showUI) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Computer rename operation completed with warnings.`n`nPlease check if the operation was successful and restart manually if needed.",
                    "Operation Completed",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }

            $success = $true # Still consider it successful as it might have worked
        }

        # Clean up files
        Remove-Item $batchFilePath -Force -ErrorAction SilentlyContinue
        Remove-Item "rename_log.txt" -Force -ErrorAction SilentlyContinue

        return $success
    }
    catch {
        $errorMsg = "Error during computer rename operation: $_"
        Add-Status $errorMsg

        if ($showUI) {
            [System.Windows.Forms.MessageBox]::Show(
                $errorMsg + "`n`nMake sure you have administrator privileges.",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }

        # Clean up files
        if (Test-Path $batchFilePath) {
            Remove-Item $batchFilePath -Force -ErrorAction SilentlyContinue
        }
        Remove-Item "rename_log.txt" -Force -ErrorAction SilentlyContinue

        return $false
    }
}

function Invoke-RenameDialog {
    Hide-MainMenu
    # Create device rename form
    $renameForm = New-Object System.Windows.Forms.Form
    $renameForm.Text = "Rename Device"
    $renameForm.Size = New-Object System.Drawing.Size(495, 470)
    $renameForm.StartPosition = "CenterScreen"
    $renameForm.BackColor = [System.Drawing.Color]::Black
    $renameForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $renameForm.MaximizeBox = $false
    $renameForm.MinimizeBox = $false

    # Add gradient background
    $renameForm.Add_Paint({
            $graphics = $_.Graphics
            $rect = New-Object System.Drawing.Rectangle(0, 0, $renameForm.Width, $renameForm.Height)
            $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
                $rect,
                [System.Drawing.Color]::FromArgb(0, 0, 0),
                [System.Drawing.Color]::FromArgb(0, 40, 0),
                [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
            )
            $graphics.FillRectangle($brush, $rect)
            $brush.Dispose()
        })

    # Create title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "RENAME DEVICE"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.Size = New-Object System.Drawing.Size(470, 40)
    $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $renameForm.Controls.Add($titleLabel)

    # Get current computer name
    $currentName = $env:COMPUTERNAME

    # Create a colored label for the current name (not $currentLabel, but $currentName itself)
    $currentNameLabel = New-Object System.Windows.Forms.Label
    $currentNameLabel.Text = $currentName
    $currentNameLabel.Font = New-Object System.Drawing.Font("Consolas", 14, [System.Drawing.FontStyle]::Bold)
    $currentNameLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 255) # Cyan color
    $currentNameLabel.BackColor = [System.Drawing.Color]::Transparent
    $currentNameLabel.AutoSize = $true
    $currentNameLabel.Location = New-Object System.Drawing.Point(180, 68)
    $renameForm.Controls.Add($currentNameLabel)

    # Current device name label
    $currentLabel = New-Object System.Windows.Forms.Label
    $currentLabel.Text = "Current Device Name:"
    $currentLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $currentLabel.ForeColor = [System.Drawing.Color]::White
    $currentLabel.Size = New-Object System.Drawing.Size(480, 30)
    $currentLabel.Location = New-Object System.Drawing.Point(10, 70)
    $currentLabel.BackColor = [System.Drawing.Color]::Transparent
    $renameForm.Controls.Add($currentLabel)

    # Device type selection group box
    $deviceGroupBox = New-Object System.Windows.Forms.GroupBox
    $deviceGroupBox.Text = "Device Type"
    $deviceGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $deviceGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 0)
    $deviceGroupBox.Size = New-Object System.Drawing.Size(460, 80)
    $deviceGroupBox.Location = New-Object System.Drawing.Point(10, 110)
    $deviceGroupBox.BackColor = [System.Drawing.Color]::Transparent

    # Desktop radio button
    $radioDesktop = New-Object System.Windows.Forms.RadioButton
    $radioDesktop.Text = "Desktop"
    $radioDesktop.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioDesktop.ForeColor = [System.Drawing.Color]::White
    $radioDesktop.Location = New-Object System.Drawing.Point(20, 30)
    $radioDesktop.Size = New-Object System.Drawing.Size(150, 30)
    $radioDesktop.BackColor = [System.Drawing.Color]::Transparent
    $radioDesktop.Checked = $true # Default selection

    # Laptop radio button
    $radioLaptop = New-Object System.Windows.Forms.RadioButton
    $radioLaptop.Text = "Laptop"
    $radioLaptop.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioLaptop.ForeColor = [System.Drawing.Color]::White
    $radioLaptop.Location = New-Object System.Drawing.Point(190, 30)
    $radioLaptop.Size = New-Object System.Drawing.Size(100, 30)
    $radioLaptop.BackColor = [System.Drawing.Color]::Transparent

    # Custom radio button
    $radioCustom = New-Object System.Windows.Forms.RadioButton
    $radioCustom.Text = "Custom"
    $radioCustom.Location = New-Object System.Drawing.Point(340, 30)
    $radioCustom.Size = New-Object System.Drawing.Size(150, 30)
    $radioCustom.ForeColor = [System.Drawing.Color]::White
    $radioCustom.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $radioCustom.BackColor = [System.Drawing.Color]::Transparent

    # Add radio buttons to group box
    $deviceGroupBox.Controls.Add($radioDesktop)
    $deviceGroupBox.Controls.Add($radioLaptop)
    $deviceGroupBox.Controls.Add($radioCustom)
    $renameForm.Controls.Add($deviceGroupBox)

    # New name label
    $newNameLabel = New-Object System.Windows.Forms.Label
    $newNameLabel.Text = "New Device Name:"
    $newNameLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $newNameLabel.ForeColor = [System.Drawing.Color]::White
    $newNameLabel.Size = New-Object System.Drawing.Size(150, 30)
    $newNameLabel.Location = New-Object System.Drawing.Point(10, 205)
    $newNameLabel.BackColor = [System.Drawing.Color]::Transparent
    $renameForm.Controls.Add($newNameLabel)

    # New name textbox
    $newNameTextBox = New-Object System.Windows.Forms.TextBox
    $newNameTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $newNameTextBox.Size = New-Object System.Drawing.Size(290, 30)
    $newNameTextBox.Location = New-Object System.Drawing.Point(180, 200)
    $newNameTextBox.BackColor = [System.Drawing.Color]::White
    $newNameTextBox.ForeColor = [System.Drawing.Color]::Black
    $newNameTextBox.Text = "HOD" # Default to Desktop
    $renameForm.Controls.Add($newNameTextBox)

    # Event handlers for radio buttons to update the default name
    $radioDesktop.Add_CheckedChanged({
            if ($radioDesktop.Checked) {
                $newNameTextBox.Text = "HOD"
            }
        })

    $radioLaptop.Add_CheckedChanged({
            if ($radioLaptop.Checked) {
                $newNameTextBox.Text = "HOL"
            }
        })

    $radioCustom.Add_CheckedChanged({
            if ($radioCustom.Checked) {
                $newNameTextBox.Text = ""
            }
        })

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(10, 300)
    $statusTextBox.Size = New-Object System.Drawing.Size(460, 120)
    $statusTextBox.BackColor = [System.Drawing.Color]::Black
    $statusTextBox.ForeColor = [System.Drawing.Color]::Lime
    $statusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $statusTextBox.ReadOnly = $true
    $statusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $statusTextBox.Text = "Ready to rename device..."
    $renameForm.Controls.Add($statusTextBox)

    # Function to add status message
    function Add-Status {
        param([string]$message)
        $statusTextBox.AppendText("`r`n$(Get-Date -Format 'HH:mm:ss') - $message")
        $statusTextBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Rename button
    $renameButton = New-Object System.Windows.Forms.Button
    $renameButton.Text = "Rename Device"
    $renameButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $renameButton.ForeColor = [System.Drawing.Color]::White
    $renameButton.BackColor = [System.Drawing.Color]::FromArgb(0, 180, 0)
    $renameButton.Size = New-Object System.Drawing.Size(200, 40)
    $renameButton.Location = New-Object System.Drawing.Point(30, 240)
    $renameButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $renameButton.Add_Click({
            $newName = $newNameTextBox.Text.Trim()

            # Disable the rename button to prevent multiple clicks
            $renameButton.Enabled = $false

            # Call the rename function with a status callback
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'result')]
            $result = Rename-DeviceWithBatch -newName $newName -statusCallback ${function:Add-Status} -showUI $true

            # Re-enable the rename button
            $renameButton.Enabled = $true
        })
    $renameForm.Controls.Add($renameButton)

    # Cancel button
    $cancelButton = New-DynamicButton  -text "Cancel" -x 250 -y 240 -width 200 -height 40 -clickAction {
        $renameForm.Close()
    } -normalColor ([System.Drawing.Color]::FromArgb(200, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(255, 50, 50)) -pressColor ([System.Drawing.Color]::FromArgb(150, 0, 0))
    $renameForm.Controls.Add($cancelButton)

    # Set the accept button (Enter key)
    $renameForm.AcceptButton = $renameButton
    # Set the cancel button (Escape key)
    $renameForm.CancelButton = $cancelButton

    # When the form is closed, show the main menu again
    $renameForm.Add_FormClosed({
            Show-MainMenu
        })

    # Show the form
    $renameForm.ShowDialog()
}

# [8] Password Functions
function Show-SetPasswordForm {
    param(
        [string]$currentUser
    )

    # Set Password Form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Set Password"
    $form.Size = New-Object System.Drawing.Size(500, 270)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::Black
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    # Title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Set Password for Current User"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.Size = New-Object System.Drawing.Size(480, 40)
    $titleLabel.Location = New-Object System.Drawing.Point(10, 20)
    $form.Controls.Add($titleLabel)

    # User label
    $userLabel = New-Object System.Windows.Forms.Label
    $userLabel.Text = "Current User:            $currentUser"
    $userLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $userLabel.ForeColor = [System.Drawing.Color]::White
    $userLabel.Size = New-Object System.Drawing.Size(480, 30)
    $userLabel.Location = New-Object System.Drawing.Point(30, 70)
    $form.Controls.Add($userLabel)

    # Password label
    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Text = "New Password:"
    $passwordLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $passwordLabel.ForeColor = [System.Drawing.Color]::White
    $passwordLabel.Size = New-Object System.Drawing.Size(130, 30)
    $passwordLabel.Location = New-Object System.Drawing.Point(30, 110)
    $form.Controls.Add($passwordLabel)

    # Password textbox
    $passwordTextBox = New-Object System.Windows.Forms.TextBox
    $passwordTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $passwordTextBox.Size = New-Object System.Drawing.Size(200, 30)
    $passwordTextBox.Location = New-Object System.Drawing.Point(180, 110)
    $passwordTextBox.BackColor = [System.Drawing.Color]::White
    $passwordTextBox.ForeColor = [System.Drawing.Color]::Black
    $passwordTextBox.UseSystemPasswordChar = $false # Mặc định hiển thị password
    $form.Controls.Add($passwordTextBox)

    # Show Password checkbox (default checked)
    $showPasswordCheckBox = New-Object System.Windows.Forms.CheckBox
    $showPasswordCheckBox.Text = "Show"
    $showPasswordCheckBox.Location = New-Object System.Drawing.Point(400, 115)
    $showPasswordCheckBox.Size = New-Object System.Drawing.Size(100, 20)
    $showPasswordCheckBox.ForeColor = [System.Drawing.Color]::White
    $showPasswordCheckBox.Font = New-Object System.Drawing.Font("Arial", 9)
    $showPasswordCheckBox.BackColor = [System.Drawing.Color]::Transparent
    $showPasswordCheckBox.Checked = $true
    $showPasswordCheckBox.Add_CheckedChanged({
            $passwordTextBox.UseSystemPasswordChar = -not $showPasswordCheckBox.Checked
        })
    $form.Controls.Add($showPasswordCheckBox)

    # Info label for empty password
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = "Leave the password field empty to set a blank password."
    $infoLabel.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $infoLabel.ForeColor = [System.Drawing.Color]::Red
    $infoLabel.Size = New-Object System.Drawing.Size(450, 20)
    $infoLabel.Location = New-Object System.Drawing.Point(70, 145)
    $form.Controls.Add($infoLabel)

    # Set Password button
    $setButton = New-Object System.Windows.Forms.Button
    $setButton.Text = "Set Password"
    $setButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $setButton.ForeColor = [System.Drawing.Color]::White
    $setButton.BackColor = [System.Drawing.Color]::FromArgb(0, 180, 0)
    $setButton.Size = New-Object System.Drawing.Size(200, 40)
    $setButton.Location = New-Object System.Drawing.Point(30, 180)
    $setButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $setButton.Add_Click({
            $password = $passwordTextBox.Text
            try {
                # Create a command to set the password
                if ([string]::IsNullOrEmpty($password)) {
                    $command = "net user $currentUser """""
                }
                else {
                    $command = "net user $currentUser $password"
                }
                $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command" -NoNewWindow -Wait -PassThru
                if ($process.ExitCode -eq 0) {
                    if ([string]::IsNullOrEmpty($password)) {
                        [System.Windows.Forms.MessageBox]::Show("Password has been removed. User '$currentUser' can now log in without a password.", "Password Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    }
                    else {
                        [System.Windows.Forms.MessageBox]::Show("Password has been changed.", "Password Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    }
                    $form.Close()
                }
                else {
                    throw "Failed to set password. Exit code: $($process.ExitCode)"
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error setting password: $_`n`nNote: This operation requires administrative privileges.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })
    $form.Controls.Add($setButton)

    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $cancelButton.ForeColor = [System.Drawing.Color]::White
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
    $cancelButton.Size = New-Object System.Drawing.Size(200, 40)
    $cancelButton.Location = New-Object System.Drawing.Point(250, 180)
    $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelButton.Add_Click({
            $form.Close()
        })
    $form.Controls.Add($cancelButton)

    # Set Accept/Cancel button for Enter/Esc
    $form.AcceptButton = $setButton
    $form.CancelButton = $cancelButton

    # Focus on password textbox when form shows
    $form.Add_Shown({
            $passwordTextBox.Focus()
        })

    # Show the form
    $form.ShowDialog()
}

function Set-UserPassword {
    param(
        [string]$user,
        [string]$password
    )
    try {
        if ([string]::IsNullOrEmpty($password)) {
            # Xóa mật khẩu (blank)
            $command = "net user $user """""
        }
        else {
            $command = "net user $user $password"
        }
        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command" -NoNewWindow -Wait -PassThru
        return $process.ExitCode -eq 0
    }
    catch {
        return $false
    }
}

function Remove-UserPassword {
    param(
        [string]$user
    )
    try {
        $command = "net user $user """""
        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command" -NoNewWindow -Wait -PassThru
        return $process.ExitCode -eq 0
    }
    catch {
        return $false
    }
}

function Invoke-SetPasswordDialog {
    $currentUser = $env:USERNAME
    Hide-MainMenu
    $result = Show-SetPasswordForm -currentUser $currentUser

    if ($result.Action -eq "set") {
        $success = Set-UserPassword -user $currentUser -password $result.Password
        if ($success) {
            if ([string]::IsNullOrEmpty($result.Password)) {
                [System.Windows.Forms.MessageBox]::Show("Password has been removed. User '$currentUser' can now log in without a password.", "Password Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Password has been changed.", "Password Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Error setting password. This operation may require administrative privileges.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    elseif ($result.Action -eq "remove") {
        $success = Remove-UserPassword -user $currentUser
        if ($success) {
            [System.Windows.Forms.MessageBox]::Show("Password has been removed. User '$currentUser' can now log in without a password.", "Password Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Error removing password. This operation may require administrative privileges.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    Show-MainMenu
}

# --- TẠO MENU 2 CỘT, TỰ ĐỘNG CO GIÃN ---
$menuButtons = @(
    @{text = '[1] Run All'; action = { [System.Windows.Forms.MessageBox]::Show('Run All!') } },
    @{text = '[6] Features'; action = { [System.Windows.Forms.MessageBox]::Show('Turn On Features!') } },
    @{text = '[2] Software'; action = { [System.Windows.Forms.MessageBox]::Show('Install All Software!') } },
    @{text = '[7] Rename'; action = { Invoke-RenameDialog } },
    @{text = '[3] Power'; action = { [System.Windows.Forms.MessageBox]::Show('Power Options!') } },
    @{text = '[8] Password'; action = { Invoke-SetPasswordDialog } },
    @{text = '[4] Volume'; action = { Invoke-VolumeManagementDialog } },
    @{text = '[9] Domain'; action = { [System.Windows.Forms.MessageBox]::Show('Join Domain!') } },
    @{text = '[5] Activate'; action = { [System.Windows.Forms.MessageBox]::Show('Activate!') } },
    @{text = '[0] Exit'; action = { $script:form.Close() } }
)

# Cấu hình các nút menu
$buttonHeight = 60
$buttonSpacingY = 10
$buttonTop = 80
$buttonLeft = 30
$buttonControls = @()

# Tạo các nút menu
for ($i = 0; $i -lt $menuButtons.Count; $i += 2) {
    # Nút bên trái
    if ($menuButtons[$i].text -eq '[0] Exit') {
        $btnL = New-DynamicButton -text $menuButtons[$i].text -x $buttonLeft -y ($buttonTop + [math]::Floor($i / 2) * ($buttonHeight + $buttonSpacingY)) -width 1 -height $buttonHeight -clickAction $menuButtons[$i].action -normalColor ([System.Drawing.Color]::FromArgb(200, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(255, 50, 50)) -pressColor ([System.Drawing.Color]::FromArgb(150, 0, 0))
    }
    else {
        $btnL = New-DynamicButton -text $menuButtons[$i].text -x $buttonLeft -y ($buttonTop + [math]::Floor($i / 2) * ($buttonHeight + $buttonSpacingY)) -width 1 -height $buttonHeight -clickAction $menuButtons[$i].action
    }
    $btnL.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $script:form.Controls.Add($btnL)
    $buttonControls += $btnL
    # Nút bên phải
    if ($i + 1 -lt $menuButtons.Count) {
        if ($menuButtons[$i + 1].text -eq '[0] Exit') {
            $btnR = New-DynamicButton -text $menuButtons[$i + 1].text -x 0 -y ($buttonTop + [math]::Floor($i / 2) * ($buttonHeight + $buttonSpacingY)) -width 1 -height $buttonHeight -clickAction $menuButtons[$i + 1].action -normalColor ([System.Drawing.Color]::FromArgb(200, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(255, 50, 50)) -pressColor ([System.Drawing.Color]::FromArgb(150, 0, 0))
        }
        else {
            $btnR = New-DynamicButton -text $menuButtons[$i + 1].text -x 0 -y ($buttonTop + [math]::Floor($i / 2) * ($buttonHeight + $buttonSpacingY)) -width 1 -height $buttonHeight -clickAction $menuButtons[$i + 1].action
        }
        $btnR.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        if ($menuButtons[$i + 1].text -eq '[7] Rename' -or $menuButtons[$i + 1].text -eq '[8] Password') {
            $btnR.Visible = $false
        }
        $script:form.Controls.Add($btnR)
        $buttonControls += $btnR
    }
}

# Cập nhật lại bố cục menu
function Update-MenuLayout {
    $formWidth = $script:form.ClientSize.Width
    $formHeight = $script:form.ClientSize.Height
    $numRows = [math]::Ceiling($buttonControls.Count / 2)
    $minBtnWidth = 120
    $minBtnHeight = 40
    $colWidth = [math]::Max($minBtnWidth, [math]::Floor(($formWidth - 3 * $buttonLeft) / 2))
    $rowHeight = [math]::Max($minBtnHeight, [math]::Floor(($formHeight - $buttonTop - 30 - ($numRows - 1) * $buttonSpacingY) / $numRows))
    for ($i = 0; $i -lt $buttonControls.Count; $i += 2) {
        $rowIdx = [math]::Floor($i / 2)
        $y = $buttonTop + $rowIdx * ($rowHeight + $buttonSpacingY)
        $buttonControls[$i].Width = $colWidth
        $buttonControls[$i].Height = $rowHeight
        $buttonControls[$i].Left = $buttonLeft
        $buttonControls[$i].Top = $y
        if ($i + 1 -lt $buttonControls.Count) {
            $buttonControls[$i + 1].Width = $colWidth
            $buttonControls[$i + 1].Height = $rowHeight
            $buttonControls[$i + 1].Left = 2 * $buttonLeft + $colWidth
            $buttonControls[$i + 1].Top = $y
        }
    }
}

# Thêm sự kiện resize
$script:form.Add_Resize({ Update-MenuLayout })
Update-MenuLayout

# Bắt đầu chương trình
$script:form.ShowDialog()