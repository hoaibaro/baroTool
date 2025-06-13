# ==============================================================================
# BAROPROVIP - VOLUME MANAGEMENT TOOL (FULLY ORGANIZED VERSION)
# ==============================================================================

# SECTION 1: ADMIN PRIVILEGES CHECK & INITIALIZATION
    Write-Warning "This script requires administrative privileges. Attempting to restart with elevation..."
    Start-Sleep -Seconds 1

    # Restart script with admin privileges
    $scriptPath = $MyInvocation.MyCommand.Path
    $arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""

    Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs

    # Exit the current non-elevated instance
    exit
}

# Hide PowerShell console window
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0) # 0 = hide

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# SECTION 2: UTILITY FUNCTIONS - Các hàm tiện ích chung
# Functions to hide/show the main menu
function Hide-MainMenu {
    $script:form.Hide()
}

function Show-MainMenu {
    $script:form.Show()
    $script:form.BringToFront()
}

# SECTION 3: UI CREATION FUNCTIONS - Các hàm tạo giao diện
# Function to create a button with green background
function New-GreenButton {
    param (
        [string]$text,
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height,
        [scriptblock]$clickAction
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $text
    $button.Location = New-Object System.Drawing.Point($x, $y)
    $button.Size = New-Object System.Drawing.Size($width, $height)
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.BackColor = [System.Drawing.Color]::FromArgb(0, 128, 0) # Dark Green
    $button.ForeColor = [System.Drawing.Color]::White
    $button.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $button.Add_Click($clickAction)

    return $button
}

# Function to create a button with red background
function New-RedButton {
    param (
        [string]$text,
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height,
        [scriptblock]$clickAction
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $text
    $button.Location = New-Object System.Drawing.Point($x, $y)
    $button.Size = New-Object System.Drawing.Size($width, $height)
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.BackColor = [System.Drawing.Color]::FromArgb(180, 0, 0) # Dark Red
    $button.ForeColor = [System.Drawing.Color]::White
    $button.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $button.Add_Click($clickAction)

    return $button
}

# Function to create a dynamic button with rounded corners and hover effects
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

    # Add rounded corners using Region
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $radius = 15 # Adjust this value to change the roundness
    $path.AddArc($button.ClientRectangle.X, $button.ClientRectangle.Y, $radius * 2, $radius * 2, 180, 90)
    $path.AddArc($button.ClientRectangle.Width - $radius * 2, $button.ClientRectangle.Y, $radius * 2, $radius * 2, 270, 90)
    $path.AddArc($button.ClientRectangle.Width - $radius * 2, $button.ClientRectangle.Height - $radius * 2, $radius * 2, $radius * 2, 0, 90)
    $path.AddArc($button.ClientRectangle.X, $button.ClientRectangle.Height - $radius * 2, $radius * 2, $radius * 2, 90, 90)
    $path.CloseAllFigures()
    $button.Region = New-Object System.Drawing.Region($path)

    # Mouse enter event - simple hover effect (no border)
    $button.Add_MouseEnter({
        # No border, just color change handled by FlatAppearance.MouseOverBackColor
        $this.FlatAppearance.BorderSize = 0
    })

    # Mouse leave event
    $button.Add_MouseLeave({
        # Keep border size at 0
        $this.FlatAppearance.BorderSize = 0
    })

    # Mouse down event
    $button.Add_MouseDown({
        # No border on press, just color change handled by FlatAppearance.MouseDownBackColor
        $this.FlatAppearance.BorderSize = 0
    })

    # Mouse up event
    $button.Add_MouseUp({
        # Keep border size at 0
        $this.FlatAppearance.BorderSize = 0
    })

    # Click event
    $button.Add_Click($clickAction)

    return $button
}

# SECTION 4: RENAME VOLUME FUNCTIONS - Các hàm đổi tên ổ đĩa
# Hàm đổi tên ổ đĩa bằng WMI
function Rename-DriveWithWMI {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DriveLetter,

        [Parameter(Mandatory = $true)]
        [string]$NewLabel
    )

    # Phương pháp 1: Sử dụng Win32_Volume
    $volume = Get-WmiObject -Query "SELECT * FROM Win32_Volume WHERE DriveLetter='$DriveLetter`:'"
    if ($volume) {
        # Đặt tên nhãn trực tiếp, đảm bảo không có khoảng trắng ở đầu và cuối
        $volume.Label = $NewLabel.Trim()
        $result = $volume.Put()
        if ($result.ReturnValue -eq 0) {
            return $true
        }
    }

    # Phương pháp 2: Sử dụng Win32_LogicalDisk
    $logicalDisk = Get-WmiObject -Query "SELECT * FROM Win32_LogicalDisk WHERE DeviceID='$DriveLetter`:'"
    if ($logicalDisk) {
        # Đặt tên nhãn trực tiếp, đảm bảo không có khoảng trắng ở đầu và cuối
        $logicalDisk.VolumeName = $NewLabel.Trim()
        $result = $logicalDisk.Put()
        if ($result.ReturnValue -eq 0) {
            return $true
        }
    }

    return $false
}

# Hàm đổi tên ổ đĩa bằng lệnh label
function Rename-DriveWithLabel {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DriveLetter,

        [Parameter(Mandatory = $true)]
        [string]$NewLabel
    )

    # Sử dụng PowerShell trực tiếp để đổi tên ổ đĩa
    # Đảm bảo không có khoảng trắng ở đầu và cuối tên nhãn
    $trimmedLabel = $NewLabel.Trim()

    # Tạo script PowerShell để đổi tên ổ đĩa
    $psScript = @"
    `$volume = Get-WmiObject -Query "SELECT * FROM Win32_Volume WHERE DriveLetter='$DriveLetter`:'"
    if (`$volume) {
        `$volume.Label = '$trimmedLabel'
        `$result = `$volume.Put()
        exit `$result.ReturnValue
    } else {
        `$logicalDisk = Get-WmiObject -Query "SELECT * FROM Win32_LogicalDisk WHERE DeviceID='$DriveLetter`:'"
        if (`$logicalDisk) {
            `$logicalDisk.VolumeName = '$trimmedLabel'
            `$result = `$logicalDisk.Put()
            exit `$result.ReturnValue
        } else {
            exit 1
        }
    }
"@

    # Lưu script vào file tạm
    $tempScriptPath = [System.IO.Path]::GetTempFileName() + ".ps1"
    Set-Content -Path $tempScriptPath -Value $psScript -Force

    # Chạy script với quyền admin
    $labelProcess = Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$tempScriptPath`"" -WindowStyle Hidden -PassThru -Wait

    # Xóa file tạm
    if (Test-Path $tempScriptPath) {
        Remove-Item $tempScriptPath -Force -ErrorAction SilentlyContinue
    }

    return ($labelProcess.ExitCode -eq 0)
}

# Tạo tiêu đề đổi tên ổ đĩa
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

# Tạo GroupBox chứa các controls cho Rename Volume
function New-RenameVolumeGroupBox {
    param([System.Windows.Forms.Panel]$parentPanel, [System.Windows.Forms.ListBox]$driveListBox)
    
    # Create GroupBox for centered content
    $renameGroupBox = New-Object System.Windows.Forms.GroupBox
    $renameGroupBox.Text = "Volume Rename Configuration"  # ✅ Thêm Text để event handler có thể tìm thấy
    $renameGroupBox.Location = New-Object System.Drawing.Point(180, 60)
    $renameGroupBox.Size = New-Object System.Drawing.Size(400, 150)
    $renameGroupBox.ForeColor = [System.Drawing.Color]::Lime
    $renameGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $parentPanel.Controls.Add($renameGroupBox)

    # Drive letter label
    $driveLetterLabel = New-Object System.Windows.Forms.Label
    $driveLetterLabel.Text = "Select Drive Letter to Rename:"
    $driveLetterLabel.Location = New-Object System.Drawing.Point(20, 30)
    $driveLetterLabel.Size = New-Object System.Drawing.Size(200, 20)
    $driveLetterLabel.ForeColor = [System.Drawing.Color]::White
    $driveLetterLabel.Font = New-Object System.Drawing.Font("Arial", 10)
    $renameGroupBox.Controls.Add($driveLetterLabel)

    # Drive letter textbox
    $driveLetterTextBox = New-Object System.Windows.Forms.TextBox
    $driveLetterTextBox.Location = New-Object System.Drawing.Point(230, 30)
    $driveLetterTextBox.Size = New-Object System.Drawing.Size(50, 20)
    $driveLetterTextBox.BackColor = [System.Drawing.Color]::Black
    $driveLetterTextBox.ForeColor = [System.Drawing.Color]::Lime
    $driveLetterTextBox.Font = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
    $driveLetterTextBox.MaxLength = 1
    $driveLetterTextBox.ReadOnly = $true
    $driveLetterTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
    $renameGroupBox.Controls.Add($driveLetterTextBox)

    # New label label
    $newLabelLabel = New-Object System.Windows.Forms.Label
    $newLabelLabel.Text = "New Volume Label:"
    $newLabelLabel.Location = New-Object System.Drawing.Point(20, 60)
    $newLabelLabel.Size = New-Object System.Drawing.Size(200, 20)
    $newLabelLabel.ForeColor = [System.Drawing.Color]::White
    $newLabelLabel.Font = New-Object System.Drawing.Font("Arial", 10)
    $renameGroupBox.Controls.Add($newLabelLabel)

    # New label textbox with script scope
    $script:newLabelTextBox = New-Object System.Windows.Forms.TextBox
    $script:newLabelTextBox.Location = New-Object System.Drawing.Point(230, 60)
    $script:newLabelTextBox.Size = New-Object System.Drawing.Size(150, 20)
    $script:newLabelTextBox.BackColor = [System.Drawing.Color]::Black
    $script:newLabelTextBox.ForeColor = [System.Drawing.Color]::Lime
    $script:newLabelTextBox.Font = New-Object System.Drawing.Font("Consolas", 11)
    $script:newLabelTextBox.Text = ""
    $renameGroupBox.Controls.Add($script:newLabelTextBox)

    # Initialize drive letter from selected drive
    if ($driveListBox.SelectedItem) {
        $selectedDrive = $driveListBox.SelectedItem.ToString()
        $driveLetter = $selectedDrive.Substring(0, 1)
        $driveLetterTextBox.Text = $driveLetter
    }
    
    return @{
        GroupBox = $renameGroupBox
        DriveLetterTextBox = $driveLetterTextBox
        NewLabelTextBox = $script:newLabelTextBox
    }
}

# Tạo nút đổi tên ổ đĩa trong GroupBox
function New-RenameActionButton {
    param([System.Windows.Forms.GroupBox]$groupBox, [System.Windows.Forms.ListBox]$driveListBox)
    
    # Rename button (inside GroupBox)
    $renameButton = New-DynamicButton -text "Rename Volume" -x 100 -y 100 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Get input values
        if (-not $driveListBox.SelectedItem) {
            Add-Status "Error: Please select a drive from the list first."
            return
        }

        $selectedDrive = $driveListBox.SelectedItem.ToString()
        $driveLetter = $selectedDrive.Substring(0, 1).ToUpper()
        $newLabel = $script:newLabelTextBox.Text.Trim()

        # Validate input
        if (-not (Test-RenameVolumeInput -driveListBox $driveListBox -newLabel $newLabel)) {
            return
        }

        # Perform rename operation
        Add-Status "Renaming drive $driveLetter to $newLabel..."
        Invoke-VolumeRenameOperation -driveLetter $driveLetter -newLabel $newLabel
    }
    $groupBox.Controls.Add($renameButton)

    # Add Enter key event handler to textbox
    $script:newLabelTextBox.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            $renameButton.PerformClick()
        }
    })
    
    return $renameButton
}

# Kiểm tra đầu vào đổi tên ổ đĩa
function Test-RenameVolumeInput {
    param(
        [System.Windows.Forms.ListBox]$driveListBox,
        [string]$newLabel
    )
    
    if (-not $driveListBox.SelectedItem) {
        Add-Status "Error: Please select a drive from the list first."
        return $false
    }

    if ($newLabel -eq "") {
        Add-Status "Error: Please enter a new label."
        return $false
    }
    
    return $true
}

# Thực hiện thao tác đổi tên ổ đĩa
function Invoke-VolumeRenameOperation {
    param(
        [string]$driveLetter,
        [string]$newLabel
    )
    
    $success = $false
    
    # Method 1: Try Set-Volume (Windows 8+)
    try {
        if (Get-Command Set-Volume -ErrorAction SilentlyContinue) {
            Set-Volume -DriveLetter $driveLetter -NewFileSystemLabel $newLabel -ErrorAction Stop
            Add-Status "Successfully renamed drive $driveLetter to $newLabel."
            $success = $true
        }
    }
    catch {
        # Silent fallback to next method
    }
    
    # Method 2: Try WMI method
    if (-not $success -and (Rename-DriveWithWMI -DriveLetter $driveLetter -NewLabel $newLabel)) {
        Add-Status "Successfully renamed drive $driveLetter to $newLabel."
        $success = $true
    }
    
    # Method 3: Try label command
    if (-not $success -and (Rename-DriveWithLabel -DriveLetter $driveLetter -NewLabel $newLabel)) {
        Add-Status "Successfully renamed drive $driveLetter to $newLabel."
        $success = $true
    }
    
    # Handle results
    if ($success) {
        $driveCount = Update-DriveList
        Add-Status "Drive list updated. Found $driveCount drives."
        $script:newLabelTextBox.Text = ""
    } else {
        Add-Status "Failed to rename drive $driveLetter. Please try renaming manually to '$newLabel'."
    }
    
    return $success
}

# SECTION 5: EXTEND VOLUME FUNCTIONS - Các hàm mở rộng ổ đĩa
# Tạo tiêu đề cho Extend Volume
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

# Tạo GroupBox chứa các controls cho Extend Volume
function New-ExtendVolumeGroupBox {
    param([System.Windows.Forms.Panel]$parentPanel)
    
    # Create GroupBox for centered content
    $extendGroupBox = New-Object System.Windows.Forms.GroupBox
    $extendGroupBox.Text = "Volume Merge Configuration"  # ✅ Thêm Text để event handler có thể tìm thấy
    $extendGroupBox.Location = New-Object System.Drawing.Point(180, 60)
    $extendGroupBox.Size = New-Object System.Drawing.Size(400, 180)
    $extendGroupBox.ForeColor = [System.Drawing.Color]::Lime
    $extendGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $parentPanel.Controls.Add($extendGroupBox)

    # Source drive label
    $sourceDriveLabel = New-Object System.Windows.Forms.Label
    $sourceDriveLabel.Text = "Source Drive (to delete):"
    $sourceDriveLabel.Location = New-Object System.Drawing.Point(75, 35)
    $sourceDriveLabel.Size = New-Object System.Drawing.Size(180, 20)
    $sourceDriveLabel.ForeColor = [System.Drawing.Color]::White
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
        GroupBox = $extendGroupBox
        SourceDriveTextBox = $script:extendSourceDriveTextBox
        TargetDriveTextBox = $script:extendTargetDriveTextBox
    }
}

# Tạo nút Extend trong GroupBox
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

# Validate input cho Extend Volume
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

# Thực hiện merge operation
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

# SECTION 6: SHRINK VOLUME FUNCTIONS - Các hàm chia nhỏ ổ đĩa
# (Shrink Volume functions would be added here if needed - they are currently implemented inline in the main application)

# SECTION 7: CHANGE DRIVE LETTER FUNCTIONS - Các hàm đổi ký tự ổ đĩa
# (Change Drive Letter functions would be added here if needed - they are currently implemented inline in the main application)

# SECTION 8: MAIN APPLICATION - Ứng dụng chính
# Create main form
$script:form = New-Object System.Windows.Forms.Form
$script:form.Text = "BAOPROVIP - SYSTEM MANAGEMENT"
$script:form.Size = New-Object System.Drawing.Size(850, 600)
$script:form.StartPosition = "CenterScreen"
$script:form.BackColor = [System.Drawing.Color]::Black
$script:form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$script:form.MaximizeBox = $false

# Add a gradient background to main form
$script:form.Paint = {
    $graphics = $_.Graphics
    $rect = New-Object System.Drawing.Rectangle(0, 0, $script:form.Width, $script:form.Height)
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect,
        [System.Drawing.Color]::FromArgb(0, 0, 0), # Black at top
        [System.Drawing.Color]::FromArgb(0, 30, 0), # Dark green at bottom
        [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
    )
    $graphics.FillRectangle($brush, $rect)
    $brush.Dispose()
}

# Tạo tiêu đề với hiệu ứng động
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "WELCOME TO BAOPROVIP"
$titleLabel.Font = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0) # Green color
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$titleLabel.Size = New-Object System.Drawing.Size($script:form.ClientSize.Width, 60)
$titleLabel.Location = New-Object System.Drawing.Point(0, 20)
$titleLabel.BackColor = [System.Drawing.Color]::Transparent
$script:form.Controls.Add($titleLabel)

# Add animation to the title
$titleTimer = New-Object System.Windows.Forms.Timer
$titleTimer.Interval = 800
$titleTimer.Add_Tick({
    if ($titleLabel.ForeColor -eq [System.Drawing.Color]::FromArgb(0, 255, 0)) {
        $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 200, 0)
    }
    else {
        $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
    }
})
$titleTimer.Start()

# Create placeholder buttons for other functions
$buttonRunAll = New-DynamicButton -text "[1] Run All" -x 30 -y 100 -width 380 -height 60 -clickAction { }
$buttonInstallSoftware = New-DynamicButton -text "[2] Install Software" -x 30 -y 180 -width 380 -height 60 -clickAction { }
$buttonPowerOptions = New-DynamicButton -text "[3] Power Options" -x 30 -y 260 -width 380 -height 60 -clickAction { }

# Change / Edit Volume - MAIN FEATURE
$buttonChangeVolume = New-DynamicButton -text "[4] Change / Edit Volume" -x 30 -y 340 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Hide the main menu
    Hide-MainMenu
    # Create volume management form
    $volumeForm = New-Object System.Windows.Forms.Form
    $volumeForm.Text = "Volume Management"
    $volumeForm.Size = New-Object System.Drawing.Size(820, 600) # Increase the size of the form
    $volumeForm.StartPosition = "CenterScreen"
    $volumeForm.BackColor = [System.Drawing.Color]::Black
    $volumeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $volumeForm.MaximizeBox = $false
    $volumeForm.MinimizeBox = $false

    # Add a gradient background
    $volumeForm.Paint = {
        $graphics = $_.Graphics
        $rect = New-Object System.Drawing.Rectangle(0, 0, $volumeForm.Width, $volumeForm.Height)
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $rect,
            [System.Drawing.Color]::FromArgb(0, 0, 0), # Black at top
            [System.Drawing.Color]::FromArgb(0, 40, 0), # Dark green at bottom
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
        )
        $graphics.FillRectangle($brush, $rect)
        $brush.Dispose()
    }

    # Title label with animation
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "VOLUME MANAGEMENT"
    $titleLabel.Location = New-Object System.Drawing.Point(0, 10) # Move the title label down
    $titleLabel.Size = New-Object System.Drawing.Size(800, 40) # Increase the size of the title label
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

    # Drive list box
    $driveListBox = New-Object System.Windows.Forms.ListBox
    $driveListBox.Location = New-Object System.Drawing.Point(20, 50) # Move the drive list box down
    $driveListBox.Size = New-Object System.Drawing.Size(760, 100) # Increase the size of the drive list box
    $driveListBox.BackColor = [System.Drawing.Color]::Black
    $driveListBox.ForeColor = [System.Drawing.Color]::Lime
    $driveListBox.Font = New-Object System.Drawing.Font("Consolas", 10)
    $volumeForm.Controls.Add($driveListBox)

    # Content Panel for function buttons
    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Location = New-Object System.Drawing.Point(20, 200)
    $contentPanel.Size = New-Object System.Drawing.Size(760, 260)
    $contentPanel.BackColor = [System.Drawing.Color]::Black
    $contentPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $volumeForm.Controls.Add($contentPanel)

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(20, 470) # Move the status text box down
    $statusTextBox.Size = New-Object System.Drawing.Size(760, 90) # Increase the size of the status text box
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

    # Function to update drive list
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

    $driveCount = Update-DriveList

    # Add a common event handler for driveListBox to update all input fields in all buttons
    $driveListBox.Add_SelectedIndexChanged({
        if ($driveListBox.SelectedItem) {
            $selectedDrive = $driveListBox.SelectedItem.ToString()
            $driveLetter = $selectedDrive.Substring(0, 1)

            # Update for Change Letter button
            if ($contentPanel.Controls.Count -gt 0 -and $contentPanel.Controls[0].Text -eq "Change Drive Letter") {
                # Update the script scope textbox directly
                if ($script:oldLetterTextBox) {
                    $script:oldLetterTextBox.Text = $driveLetter
                }
            }

            # Update for Shrink Volume button
            if ($contentPanel.Controls.Count -gt 0 -and $contentPanel.Controls[0].Text -eq "Shrink Volume and Create New Partition") {
                # Use script scope variable directly
                if ($script:selectedDriveTextBox) {
                    $script:selectedDriveTextBox.Text = $driveLetter
                }
            }

            # Update for Extend Volume button
            if ($contentPanel.Controls.Count -gt 0 -and $contentPanel.Controls[0].Text -eq "Extend Volume by Merging") {
                # Use script scope variables directly
                if ($script:extendSourceDriveTextBox -and $script:extendTargetDriveTextBox) {
                    # If source drive is empty, fill it
                    if ($script:extendSourceDriveTextBox.Text -eq "") {
                        $script:extendSourceDriveTextBox.Text = $driveLetter
                    }
                    # Otherwise, if target drive is empty and different from source, fill it
                    elseif ($script:extendTargetDriveTextBox.Text -eq "" -and $driveLetter -ne $script:extendSourceDriveTextBox.Text) {
                        $script:extendTargetDriveTextBox.Text = $driveLetter
                    }
                }
            }

            # Update for Rename Volume button
            if ($contentPanel.Controls.Count -gt 0 -and $contentPanel.Controls[0].Text -eq "Rename Volume") {
                # Method 1: Find the GroupBox by Text property
                $renameGroupBox = $contentPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.GroupBox] -and $_.Text -eq "Volume Rename Configuration" }
                
                # Method 2: Fallback - find any GroupBox in the rename panel (if Method 1 fails)
                if (-not $renameGroupBox) {
                    $renameGroupBox = $contentPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.GroupBox] }
                }
                
                if ($renameGroupBox) {
                    # Find the drive letter textbox inside the GroupBox
                    $driveLetterTextBox = $renameGroupBox.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] -and $_.Location.X -eq 230 -and $_.Location.Y -eq 30 }

                    if ($driveLetterTextBox) {
                        $driveLetterTextBox.Text = $driveLetter
                        # Don't auto-fill the new label textbox - let user type what they want
                    }
                }
            }
        }
    })

    # Change Drive Letter button
    $btnChangeDriveLetter = New-DynamicButton -text "Change Letter" -x 20 -y 150 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
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
        $changeGroupBox.Location = New-Object System.Drawing.Point(180, 60)
        $changeGroupBox.Size = New-Object System.Drawing.Size(400, 150)
        $changeGroupBox.ForeColor = [System.Drawing.Color]::Lime
        $changeGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $contentPanel.Controls.Add($changeGroupBox)

        # Old drive letter label
        $oldLetterLabel = New-Object System.Windows.Forms.Label
        $oldLetterLabel.Text = "Select Drive Letter to Change:"
        $oldLetterLabel.Location = New-Object System.Drawing.Point(20, 30)
        $oldLetterLabel.Size = New-Object System.Drawing.Size(200, 20)
        $oldLetterLabel.ForeColor = [System.Drawing.Color]::White
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

    # Shrink Volume button
    $btnShrinkVolume = New-DynamicButton -text "Shrink Volume" -x 180 -y 150 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Clear the content panel
        $contentPanel.Controls.Clear()

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

        # Selected drive letter label
        $selectedDriveLabel = New-Object System.Windows.Forms.Label
        $selectedDriveLabel.Text = "Selected Drive Letter:"
        $selectedDriveLabel.Location = New-Object System.Drawing.Point(20, 50)
        $selectedDriveLabel.Size = New-Object System.Drawing.Size(150, 20)
        $selectedDriveLabel.ForeColor = [System.Drawing.Color]::White
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

        # Partition size options group box
        $partitionGroupBox = New-Object System.Windows.Forms.GroupBox
        $partitionGroupBox.Text = "Choose Partition Size"
        $partitionGroupBox.Location = New-Object System.Drawing.Point(20, 80)
        $partitionGroupBox.Size = New-Object System.Drawing.Size(720, 120)
        $partitionGroupBox.ForeColor = [System.Drawing.Color]::Lime
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
        $script:radio80GB.Text = "80GB (recommended for 256GB drives)"
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
            } else {
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

            # Determine partition size
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
                            return
                        }
                        if ($sizeMB -gt 2097152) { # 2TB limit
                            Add-Status "Error: Custom size cannot exceed 2,097,152 MB (2 TB)."
                            return
                        }
                    }
                    catch {
                        Add-Status "Error processing custom size: $_"
                        return
                    }
                } else {
                    Add-Status "Error: Custom size must be a valid number (digits only)."
                    return
                }
            }
            else {
                Add-Status "Error: Please select a partition size option."
                return
            }

            # Validate drive exists and get info
            try {
                $driveInfo = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DeviceID -eq "$($driveLetter):" }
                if (-not $driveInfo) {
                    Add-Status "Error: Drive $driveLetter does not exist."
                    return
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
                        return
                    }
                }
                catch {
                    # Fallback: Use 80% of free space as safe shrink limit
                    $maxShrinkMB = [math]::Floor($freeSpaceMB * 0.8)

                    if ($sizeMB -gt $maxShrinkMB) {
                        Add-Status "Error: Requested size ($sizeMB MB) exceeds estimated safe shrink limit ($maxShrinkMB MB)."
                        Add-Status "Try a smaller size or free up more space on the drive."
                        return
                    }
                }
            }
            catch {
                Add-Status "Error getting drive information: $_"
                return
            }

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

            Add-Status "Shrinking drive $driveLetter and creating new partition of $sizeMB MB..."
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
                $batchProcess.WaitForExit()

                # Read status file and display in status box
                if (Test-Path "shrink_status.txt") {
                    $statusContent = Get-Content "shrink_status.txt" -Raw
                    Remove-Item "shrink_status.txt" -Force -ErrorAction SilentlyContinue
                }

                # Check if operation was successful (using exact install.ps1 logic)
                if ($batchProcess.ExitCode -eq 0) {
                    Add-Status "Operation completed successfully."
                    Add-Status "Shrunk drive $driveLetter and created new partition."

                    # Refresh drive list
                    Start-Sleep -Seconds 2

                    # Tìm ổ đĩa mới được tạo (exact same logic as install.ps1)
                    $newDriveFound = $false
                    $newDriveLetter = ""

                    # Đợi một chút để đảm bảo hệ thống đã cập nhật
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
        $contentPanel.Controls.Add($shrinkButton)

        # Update drive letter from selected drive IMMEDIATELY after controls are added
        if ($driveListBox.SelectedItem) {
            $selectedDrive = $driveListBox.SelectedItem.ToString()
            $driveLetter = $selectedDrive.Substring(0, 1)
            $script:selectedDriveTextBox.Text = $driveLetter
        }

        Add-Status "Ready to shrink volume. Select a drive, choose partition size, then click Shrink."
    }
    $volumeForm.Controls.Add($btnShrinkVolume)

    # Rename Volume button
    $btnRenameVolume = New-DynamicButton -text "Rename Volume" -x 340 -y 150 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
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

    # Extend Volume button
    $btnExtendVolume = New-DynamicButton -text "Extend Volume" -x 500 -y 150 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
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

    # Return to Main Menu button
    $btnReturn = New-DynamicButton -text "Return" -x 660 -y 150 -width 120 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        $volumeForm.Close()
    }
    $volumeForm.Controls.Add($btnReturn)

    # When the form is closed, show the main menu again
    $volumeForm.Add_FormClosed({
        Show-MainMenu
    })

    # Show the form
    $volumeForm.ShowDialog()
}
# Activate
$buttonActivate = New-DynamicButton -text "[5] Activate" -x 30 -y 420 -width 380 -height 60 -clickAction { }
# Turn On Features
$buttonTurnOnFeatures = New-DynamicButton -text "[6] Turn On Features" -x 430 -y 100 -width 380 -height 60 -clickAction { }
# Rename Device
$buttonRenameDevice = New-DynamicButton -text "[7] Rename Device" -x 430 -y 180 -width 380 -height 60 -clickAction { }
# Set Password
$buttonSetPassword = New-DynamicButton -text "[8] Set Password" -x 430 -y 260 -width 380 -height 60 -clickAction {
    # Hide the main menu
    Hide-MainMenu
    # Create password form
    $passwordForm = New-Object System.Windows.Forms.Form
    $passwordForm.Text = "Password Management"
    $passwordForm.Size = New-Object System.Drawing.Size(500, 450)
    $passwordForm.StartPosition = "CenterScreen"
 }
# Join Domain
$buttonJoinDomain = New-DynamicButton -text "[9] Join Domain" -x 430 -y 340 -width 380 -height 60 -clickAction {
    # Hide the main menu
    Hide-MainMenu
    # Create domain/workgroup join form
    $joinForm = New-Object System.Windows.Forms.Form
    $joinForm.Text = "Domain Management"
    $joinForm.Size = New-Object System.Drawing.Size(500, 450)
    $joinForm.StartPosition = "CenterScreen"
    $joinForm.BackColor = [System.Drawing.Color]::Black
    $joinForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $joinForm.MaximizeBox = $false
    $joinForm.MinimizeBox = $false

    # Create title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Domain Management"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.Size = New-Object System.Drawing.Size(480, 40)
    $titleLabel.Location = New-Object System.Drawing.Point(10, 20)
    $joinForm.Controls.Add($titleLabel)

    # Get current computer information
    $currentName = $env:COMPUTERNAME
    try {
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem
        $currentWorkgroup = $computerSystem.Domain
        $isDomain = $computerSystem.PartOfDomain
    }
    catch {
        $currentWorkgroup = "Unknown"
        $isDomain = $false
    }

    # Current computer info label
    $currentLabel = New-Object System.Windows.Forms.Label
    $currentLabel.Text = "Current Computer Name: $currentName"
    $currentLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $currentLabel.ForeColor = [System.Drawing.Color]::White
    $currentLabel.Size = New-Object System.Drawing.Size(480, 30)
    $currentLabel.Location = New-Object System.Drawing.Point(20, 70)
    $joinForm.Controls.Add($currentLabel)

    # Current domain/workgroup info
    $domainLabel = New-Object System.Windows.Forms.Label
    if ($isDomain) {
        $domainLabel.Text = "Currently joined to DOMAIN: $currentWorkgroup"
    }
    else {
        $domainLabel.Text = "Currently joined to WORKGROUP: $currentWorkgroup"
    }
    $domainLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $domainLabel.ForeColor = [System.Drawing.Color]::White
    $domainLabel.Size = New-Object System.Drawing.Size(480, 30)
    $domainLabel.Location = New-Object System.Drawing.Point(20, 100)
    $joinForm.Controls.Add($domainLabel)

    # Create radio buttons for domain/workgroup selection
    $groupBox = New-Object System.Windows.Forms.GroupBox
    $groupBox.Text = "Select Option"
    $groupBox.Font = New-Object System.Drawing.Font("Arial", 10)
    $groupBox.ForeColor = [System.Drawing.Color]::White
    $groupBox.Size = New-Object System.Drawing.Size(460, 80)
    $groupBox.Location = New-Object System.Drawing.Point(20, 140)
    $groupBox.BackColor = [System.Drawing.Color]::Black

# Create other main menu buttons
$buttonActivate = New-DynamicButton -text "[5] Activate" -x 430 -y 100 -width 380 -height 60 -clickAction { }
$buttonTurnOnFeatures = New-DynamicButton -text "[6] Turn On Features" -x 430 -y 180 -width 380 -height 60 -clickAction { }
$buttonRenameDevice = New-DynamicButton -text "[7] Rename Device" -x 430 -y 260 -width 380 -height 60 -clickAction { }
$buttonSetPassword = New-DynamicButton -text "[8] Set Password" -x 430 -y 340 -width 380 -height 60 -clickAction { }
$buttonJoinDomain = New-DynamicButton -text "[9] Join Domain" -x 430 -y 420 -width 380 -height 60 -clickAction { }

# Exit button
$buttonExit = New-DynamicButton -text "[10] Exit" -x 30 -y 420 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
    $script:form.Close()
}

# Add buttons to form
$script:form.Controls.Add($buttonRunAll)
$script:form.Controls.Add($buttonInstallSoftware)
$script:form.Controls.Add($buttonPowerOptions)
$script:form.Controls.Add($buttonChangeVolume)
$script:form.Controls.Add($buttonActivate)
$script:form.Controls.Add($buttonTurnOnFeatures)
$script:form.Controls.Add($buttonRenameDevice)
$script:form.Controls.Add($buttonSetPassword)
$script:form.Controls.Add($buttonJoinDomain)
$script:form.Controls.Add($buttonExit)

# Display form
$script:form.ShowDialog()