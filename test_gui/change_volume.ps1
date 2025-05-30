# Check for admin privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
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

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "BAOPROVIP - SYSTEM MANAGEMENT"
$form.Size = New-Object System.Drawing.Size(850, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::Black
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Add a gradient background to main form
$form.Paint = {
    $graphics = $_.Graphics
    $rect = New-Object System.Drawing.Rectangle(0, 0, $form.Width, $form.Height)
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
$titleLabel.Size = New-Object System.Drawing.Size($form.ClientSize.Width, 60)
$titleLabel.Location = New-Object System.Drawing.Point(0, 20)
$titleLabel.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($titleLabel)

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

# Change / Edit Volume
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

    # Populate drive list when form loads
    Add-Status "Getting list of drives..."
    try {
        $driveCount = Update-DriveList
        Add-Status "Found $driveCount drives."
    }
    catch {
        Add-Status "Error getting drive list: $_"
    }

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
                # Find the source and target drive textboxes
                $sourceDriveTextBox = $contentPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] -and $_.Location.X -eq 330 }
                $targetDriveTextBox = $contentPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] -and $_.Location.X -eq 580 }

                if ($sourceDriveTextBox -and $targetDriveTextBox) {
                    # If source drive is empty, fill it
                    if ($sourceDriveTextBox.Text -eq "") {
                        $sourceDriveTextBox.Text = $driveLetter
                    }
                    # Otherwise, if target drive is empty and different from source, fill it
                    elseif ($targetDriveTextBox.Text -eq "" -and $driveLetter -ne $sourceDriveTextBox.Text) {
                        $targetDriveTextBox.Text = $driveLetter
                    }
                }
            }

            # Update for Rename Volume button
            if ($contentPanel.Controls.Count -gt 0 -and $contentPanel.Controls[0].Text -eq "Rename Volume") {
                # Find the drive letter textbox
                $driveLetterTextBox = $contentPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] -and $_.Location.X -eq 310 -and $_.Location.Y -eq 50 }
                $newLabelTextBox = $contentPanel.Controls | Where-Object { $_ -is [System.Windows.Forms.TextBox] -and $_.Location.X -eq 310 -and $_.Location.Y -eq 80 }

                if ($driveLetterTextBox) {
                    $driveLetterTextBox.Text = $driveLetter

                    # Get current volume name
                    $drives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } }, `
                    @{Name = 'VolumeName'; Expression = { $_.VolumeName } }

                    $currentVolumeName = ""
                    foreach ($drive in $drives) {
                        if ($drive.Name -eq "$($driveLetter):") {
                            $currentVolumeName = $drive.VolumeName
                            break
                        }
                    }

                    # Set current volume name as default text
                    if ($newLabelTextBox) {
                        $newLabelTextBox.Text = $currentVolumeName
                    }
                }
            }
        }
    })

    # Create buttons horizontally
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
        $changeGroupBox.Text = "Drive Letter Configuration"
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
        $script:oldLetterTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
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

        # Note: The driveListBox.SelectedIndexChanged event is now handled at the form level

        # Change button (inside GroupBox)
        $changeButton = New-DynamicButton -text "Change Drive Letter" -x 100 -y 100 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
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

                    # Refresh drive list
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
        $newLabelTextBox = New-Object System.Windows.Forms.TextBox
        $newLabelTextBox.Location = New-Object System.Drawing.Point(450, 50)
        $newLabelTextBox.Size = New-Object System.Drawing.Size(250, 25)
        $newLabelTextBox.BackColor = [System.Drawing.Color]::Black
        $newLabelTextBox.ForeColor = [System.Drawing.Color]::Lime
        $newLabelTextBox.Font = New-Object System.Drawing.Font("Consolas", 11)
        $newLabelTextBox.Text = "GAME"
        $contentPanel.Controls.Add($newLabelTextBox)

        # Shrink button
        $shrinkButton = New-DynamicButton -text "Shrink" -x 275 -y 210 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            $driveLetter = $script:selectedDriveTextBox.Text.Trim().ToUpper()
            $newLabel = $newLabelTextBox.Text.Trim()

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
                Add-Status "Selected 80GB partition (82,020 MB)."
            }
            elseif ($script:radio200GB.Checked) {
                $sizeMB = 204955
                Add-Status "Selected 200GB partition (204,955 MB)."
            }
            elseif ($script:radio500GB.Checked) {
                $sizeMB = 512000
                Add-Status "Selected 500GB partition (512,000 MB)."
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
                        Add-Status "Selected custom size: $sizeMB MB ($([math]::Round($sizeMB/1024, 1)) GB)."
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
                $usedSpaceMB = $totalSizeMB - $freeSpaceMB

                Add-Status "Drive $driveLetter info: Total: $totalSizeMB MB, Used: $usedSpaceMB MB, Free: $freeSpaceMB MB"

                # Get actual shrinkable space using PowerShell (more accurate)
                try {
                    $partition = Get-Partition -DriveLetter $driveLetter -ErrorAction Stop
                    $shrinkInfo = Get-PartitionSupportedSize -DriveLetter $driveLetter -ErrorAction Stop
                    $maxShrinkBytes = $partition.Size - $shrinkInfo.SizeMin
                    $maxShrinkMB = [math]::Floor($maxShrinkBytes / 1MB)

                    Add-Status "Maximum shrinkable space: $maxShrinkMB MB"

                    if ($sizeMB -gt $maxShrinkMB) {
                        Add-Status "Error: Requested size ($sizeMB MB) exceeds maximum shrinkable space ($maxShrinkMB MB)."
                        Add-Status "Try running disk defragmentation first or choose a smaller size."
                        return
                    }
                }
                catch {
                    Add-Status "Warning: Could not get precise shrink info. Using fallback calculation."
                    # Fallback: Use 80% of free space as safe shrink limit
                    $maxShrinkMB = [math]::Floor($freeSpaceMB * 0.8)
                    Add-Status "Estimated maximum shrinkable space: $maxShrinkMB MB"

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
                    Add-Status "---- Operation Log ----"
                    Add-Status $statusContent
                    Add-Status "---- End of Log ----"
                    Remove-Item "shrink_status.txt" -Force -ErrorAction SilentlyContinue
                }

                # Check if operation was successful (using exact install.ps1 logic)
                if ($batchProcess.ExitCode -eq 0) {
                    Add-Status "Operation completed successfully."
                    Add-Status "Shrunk drive $driveLetter and created new partition."

                    # Refresh drive list
                    Add-Status "Refreshing drive list..."
                    Start-Sleep -Seconds 2

                    # Tìm ổ đĩa mới được tạo (exact same logic as install.ps1)
                    $newDriveFound = $false
                    $newDriveLetter = ""

                    # Đợi một chút để đảm bảo hệ thống đã cập nhật
                    Start-Sleep -Seconds 2
                    Add-Status "Scanning for new drives..."

                    # Lấy danh sách ổ đĩa hiện tại
                    $currentDrives = Get-WmiObject Win32_LogicalDisk | Select-Object DeviceID, VolumeName

                    # Tìm ổ đĩa có tên "New Volume" hoặc ổ đĩa trống
                    foreach ($drive in $currentDrives) {
                        if ($drive.DeviceID -ne "$($driveLetter):" -and
                            ($drive.VolumeName -eq "New Volume" -or $drive.VolumeName -eq "")) {
                            $newDriveFound = $true
                            $newDriveLetter = $drive.DeviceID.TrimEnd(":")
                            Add-Status "Found new drive: $newDriveLetter"
                            break
                        }
                    }



                    # Nếu tìm thấy ổ đĩa mới, đổi tên nó (fix scope issue)
                    if ($newDriveFound) {
                        # Debug: Check what we have for label
                        Add-Status "Debug: newLabel variable = '$newLabel'"
                        Add-Status "Debug: newLabelTextBox.Text = '$($newLabelTextBox.Text)'"

                        # Try to get label from textbox directly first
                        try {
                            $actualNewLabel = $newLabelTextBox.Text.Trim()
                            if ([string]::IsNullOrEmpty($actualNewLabel)) {
                                $actualNewLabel = $newLabel
                            }
                            if ([string]::IsNullOrEmpty($actualNewLabel)) {
                                $actualNewLabel = "GAME"  # Fallback default
                            }
                        }
                        catch {
                            $actualNewLabel = if (-not [string]::IsNullOrEmpty($newLabel)) { $newLabel } else { "GAME" }
                        }

                        Add-Status "Debug: Final actualNewLabel = '$actualNewLabel'"
                        Add-Status "Renaming new drive $newDriveLetter to $actualNewLabel..."

                        # Use existing rename functions instead of duplicate code
                        $renameSuccess = $false

                        # Try WMI method first (from existing Rename-DriveWithWMI function)
                        try {
                            $volume = Get-WmiObject -Query "SELECT * FROM Win32_Volume WHERE DriveLetter='$newDriveLetter`:'"
                            if ($volume) {
                                $volume.Label = $actualNewLabel.Trim()
                                $result = $volume.Put()
                                if ($result.ReturnValue -eq 0) {
                                    Add-Status "Successfully renamed drive $newDriveLetter to $actualNewLabel using WMI."
                                    $renameSuccess = $true
                                }
                            }
                        }
                        catch {
                            # Continue to next method
                        }

                        # Try PowerShell Set-Volume if WMI fails
                        if (-not $renameSuccess) {
                            try {
                                if (Get-Command Set-Volume -ErrorAction SilentlyContinue) {
                                    Set-Volume -DriveLetter $newDriveLetter -NewFileSystemLabel $actualNewLabel -ErrorAction Stop
                                    Add-Status "Successfully renamed drive $newDriveLetter to $actualNewLabel using Set-Volume."
                                    $renameSuccess = $true
                                }
                            }
                            catch {
                                # Continue to next method
                            }
                        }

                        # Try label command if previous methods fail
                        if (-not $renameSuccess) {
                            try {
                                $labelProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c label $newDriveLetter`:$actualNewLabel" -WindowStyle Hidden -PassThru -Wait
                                if ($labelProcess.ExitCode -eq 0) {
                                    Add-Status "Successfully renamed drive $newDriveLetter to $actualNewLabel using label command."
                                    $renameSuccess = $true
                                }
                            }
                            catch {
                                # All methods failed
                            }
                        }

                        if (-not $renameSuccess) {
                            Add-Status "Failed to rename drive $newDriveLetter automatically. Please rename it manually to '$actualNewLabel'."
                        }

                        # Wait for rename to complete
                        Start-Sleep -Seconds 1
                    }
                    else {
                        Add-Status "Could not find the newly created drive. Please rename it manually."
                    }

                    # Cập nhật giao diện
                    Add-Status "Updating drive list..."

                    # Refresh drive list using the existing function
                    $driveCount = Update-DriveList
                    Add-Status "Drive list updated successfully. Found $driveCount drives."
                }
                else {
                    Add-Status "Operation completed with warnings. Check the log above for details."
                    Add-Status "If the partition was created but without label, you can rename it manually."
                }

                # Clean up batch file
                Remove-Item $batchFilePath -Force -ErrorAction SilentlyContinue
            }
            catch {
                Add-Status "Error running batch operation: $_"
                Add-Status "Make sure you have administrator privileges."

                # Clean up batch file
                Remove-Item $batchFilePath -Force -ErrorAction SilentlyContinue
            }
        }
        $contentPanel.Controls.Add($shrinkButton)

        # Cancel button
        $cancelButton = New-DynamicButton -text "Cancel" -x 485 -y 210 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
            $contentPanel.Controls.Clear()
        }
        $contentPanel.Controls.Add($cancelButton)

        # Update drive letter from selected drive IMMEDIATELY after controls are added
        if ($driveListBox.SelectedItem) {
            $selectedDrive = $driveListBox.SelectedItem.ToString()
            $driveLetter = $selectedDrive.Substring(0, 1)
            $script:selectedDriveTextBox.Text = $driveLetter
            Add-Status "Selected drive initialized to: $driveLetter"
        }

        Add-Status "Ready to shrink volume. Select a drive from the list, choose partition size, then click Shrink."
    }
    $volumeForm.Controls.Add($btnShrinkVolume)

    # Extend Volume button
    $btnExtendVolume = New-DynamicButton -text "Extend Volume" -x 340 -y 150 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Clear the content panel
        $contentPanel.Controls.Clear()

        # Title label
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "Extend Volume by Merging"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 10)
        $titleLabel.Size = New-Object System.Drawing.Size(760, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $contentPanel.Controls.Add($titleLabel)

        # Source drive label
        $sourceDriveLabel = New-Object System.Windows.Forms.Label
        $sourceDriveLabel.Text = "Source Drive (to delete):"
        $sourceDriveLabel.Location = New-Object System.Drawing.Point(150, 70)
        $sourceDriveLabel.Size = New-Object System.Drawing.Size(180, 20)
        $sourceDriveLabel.ForeColor = [System.Drawing.Color]::White
        $sourceDriveLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $contentPanel.Controls.Add($sourceDriveLabel)

        # Source drive textbox
        $sourceDriveTextBox = New-Object System.Windows.Forms.TextBox
        $sourceDriveTextBox.Location = New-Object System.Drawing.Point(330, 70)
        $sourceDriveTextBox.Size = New-Object System.Drawing.Size(50, 20)
        $sourceDriveTextBox.BackColor = [System.Drawing.Color]::Black
        $sourceDriveTextBox.ForeColor = [System.Drawing.Color]::Lime
        $sourceDriveTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $sourceDriveTextBox.MaxLength = 1
        $sourceDriveTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center

        # Thêm xử lý sự kiện khi nhấn Enter
        $sourceDriveTextBox.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true  # Ngăn chặn tiếng "beep"
                $targetDriveTextBox.Focus()  # Chuyển focus đến ô target drive
            }
        })

        $contentPanel.Controls.Add($sourceDriveTextBox)

        # Target drive label
        $targetDriveLabel = New-Object System.Windows.Forms.Label
        $targetDriveLabel.Text = "Target Drive (to expand):"
        $targetDriveLabel.Location = New-Object System.Drawing.Point(400, 70)
        $targetDriveLabel.Size = New-Object System.Drawing.Size(180, 20)
        $targetDriveLabel.ForeColor = [System.Drawing.Color]::White
        $targetDriveLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $contentPanel.Controls.Add($targetDriveLabel)

        # Target drive textbox
        $targetDriveTextBox = New-Object System.Windows.Forms.TextBox
        $targetDriveTextBox.Location = New-Object System.Drawing.Point(580, 70)
        $targetDriveTextBox.Size = New-Object System.Drawing.Size(50, 20)
        $targetDriveTextBox.BackColor = [System.Drawing.Color]::Black
        $targetDriveTextBox.ForeColor = [System.Drawing.Color]::Lime
        $targetDriveTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $targetDriveTextBox.MaxLength = 1
        $targetDriveTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center

        # Thêm xử lý sự kiện khi nhấn Enter
        $targetDriveTextBox.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true  # Ngăn chặn tiếng "beep"
                $mergeButton.PerformClick()  # Kích hoạt nút Merge
            }
        })

        $contentPanel.Controls.Add($targetDriveTextBox)

        # Warning label
        $warningLabel = New-Object System.Windows.Forms.Label
        $warningLabel.Text = "WARNING: This will DELETE the source drive and all its data!"
        $warningLabel.Location = New-Object System.Drawing.Point(0, 100) # Căn giữa theo chiều ngang
        $warningLabel.Size = New-Object System.Drawing.Size(760, 20) # Sử dụng toàn bộ chiều rộng của panel
        $warningLabel.ForeColor = [System.Drawing.Color]::Red
        $warningLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $warningLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter # Căn giữa văn bản
        $contentPanel.Controls.Add($warningLabel)


        # Function to add status message to the main status box for merge operations
        function Add-MergeStatus {
            param(
                [Parameter(Mandatory=$true)]
                [string]$message,

                [Parameter(Mandatory=$false)]
                [switch]$NoNewLine,

                [Parameter(Mandatory=$false)]
                [switch]$ClearLine
            )

            # Use the main status box
            Add-Status $message
        }

        # Note: The driveListBox.SelectedIndexChanged event is now handled at the form level

        # Merge button
        $mergeButton = New-DynamicButton -text "Merge and Extend Volume" -x 150 -y 140 -width 250 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            # Lấy và kiểm tra đầu vào
            $sourceDrive = $sourceDriveTextBox.Text.Trim().ToUpper()
            $targetDrive = $targetDriveTextBox.Text.Trim().ToUpper()

            # Hiển thị thông tin đầu vào
            Add-Status "Extend Volume: Source drive entered: '$sourceDrive'"
            Add-Status "Extend Volume: Target drive entered: '$targetDrive'"

            # Kiểm tra đầu vào chi tiết
            if ([string]::IsNullOrEmpty($sourceDrive)) {
                Add-MergeStatus "Error: Please enter a source drive letter."
                return
            }

            if ([string]::IsNullOrEmpty($targetDrive)) {
                Add-MergeStatus "Error: Please enter a target drive letter."
                return
            }

            # Kiểm tra xem có phải chữ cái hợp lệ không
            if (-not ($sourceDrive -match '^[A-Z]$')) {
                Add-MergeStatus "Error: Source drive must be a single letter (A-Z)."
                return
            }

            if (-not ($targetDrive -match '^[A-Z]$')) {
                Add-MergeStatus "Error: Target drive must be a single letter (A-Z)."
                return
            }

            if ($sourceDrive -eq $targetDrive) {
                Add-MergeStatus "Error: Source and target drives cannot be the same."
                return
            }

            # Kiểm tra xem ổ đĩa có tồn tại không
            $existingDrives = Get-CimInstance -ClassName Win32_LogicalDisk | Select-Object -ExpandProperty DeviceID | ForEach-Object { $_.Substring(0, 1) }

            if ($existingDrives -notcontains $sourceDrive) {
                Add-MergeStatus "Error: Source drive $sourceDrive does not exist."
                return
            }

            if ($existingDrives -notcontains $targetDrive) {
                Add-MergeStatus "Error: Target drive $targetDrive does not exist."
                return
            }

            Add-MergeStatus "Both drives exist. Proceeding with validation..."

            # Kiểm tra xem hai ổ đĩa có nằm trên cùng một đĩa vật lý không - phương pháp tối ưu
            try {
                Add-MergeStatus "Checking if drives are on the same physical disk..."

                # Hiển thị thông tin chi tiết về ổ đĩa để debug
                Add-MergeStatus "Getting detailed disk information..."

                # Phương pháp 1: Sử dụng Get-Partition
                try {
                    $sourcePartition = Get-Partition -DriveLetter $sourceDrive -ErrorAction Stop
                    $targetPartition = Get-Partition -DriveLetter $targetDrive -ErrorAction Stop

                    $sourceDiskNumber = $sourcePartition.DiskNumber
                    $targetDiskNumber = $targetPartition.DiskNumber

                    Add-MergeStatus "Method 1 (Get-Partition):"
                    Add-MergeStatus "Source drive $sourceDrive is on disk number: $sourceDiskNumber"
                    Add-MergeStatus "Target drive $targetDrive is on disk number: $targetDiskNumber"
                }
                catch {
                    Add-MergeStatus "Method 1 failed: $_"
                    $sourceDiskNumber = $null
                    $targetDiskNumber = $null
                }

                # Phương pháp 2: Sử dụng Get-CimInstance nếu phương pháp 1 thất bại
                if ($null -eq $sourceDiskNumber -or $null -eq $targetDiskNumber) {
                    Add-MergeStatus "Trying Method 2 (CIM)..."
                    try {
                        # Lấy thông tin ổ đĩa logic (không cần lưu biến vì chúng ta chỉ cần thông tin partition)
                        # Truy vấn trực tiếp thông tin partition từ ổ đĩa logic

                        # Lấy thông tin partition
                        $sourcePartitions = Get-CimInstance -Query "ASSOCIATORS OF {Win32_LogicalDisk.DeviceID='$($sourceDrive):'} WHERE ResultClass=Win32_DiskPartition"
                        $targetPartitions = Get-CimInstance -Query "ASSOCIATORS OF {Win32_LogicalDisk.DeviceID='$($targetDrive):'} WHERE ResultClass=Win32_DiskPartition"

                        if ($sourcePartitions) {
                            $sourceDiskNumber = $sourcePartitions[0].DiskIndex
                            Add-MergeStatus "Method 2: Source drive $sourceDrive is on disk number: $sourceDiskNumber"
                        }
                        if ($targetPartitions) {
                            $targetDiskNumber = $targetPartitions[0].DiskIndex
                            Add-MergeStatus "Method 2: Target drive $targetDrive is on disk number: $targetDiskNumber"
                        }
                    }
                    catch {
                        Add-MergeStatus "Method 2 failed: $_"
                    }
                }

                # Phương pháp 3: Sử dụng diskpart để lấy thông tin
                if ($null -eq $sourceDiskNumber -or $null -eq $targetDiskNumber) {
                    Add-MergeStatus "Trying Method 3 (diskpart)..."
                    try {
                        # Tạo script diskpart để lấy thông tin
                        $diskpartScript = @"
list volume
"@
                        Set-Content -Path "get_disk_info.txt" -Value $diskpartScript

                        # Chạy diskpart
                        $diskpartOutput = & diskpart /s get_disk_info.txt

                        # Phân tích output
                        $volumeInfo = $diskpartOutput | Where-Object { $_ -match "Volume\s+\d+" }

                        foreach ($line in $volumeInfo) {
                            if ($line -match "\s+$sourceDrive\s+") {
                                if ($line -match "Disk\s+(\d+)") {
                                    $sourceDiskNumber = $matches[1]
                                    Add-MergeStatus "Method 3: Source drive $sourceDrive is on disk number: $sourceDiskNumber"
                                }
                            }
                            if ($line -match "\s+$targetDrive\s+") {
                                if ($line -match "Disk\s+(\d+)") {
                                    $targetDiskNumber = $matches[1]
                                    Add-MergeStatus "Method 3: Target drive $targetDrive is on disk number: $targetDiskNumber"
                                }
                            }
                        }

                        # Xóa file tạm
                        Remove-Item "get_disk_info.txt" -ErrorAction SilentlyContinue
                    }
                    catch {
                        Add-MergeStatus "Method 3 failed: $_"
                    }
                }

                # Hiển thị thông tin
                if ($null -ne $sourceDiskNumber) {
                    Add-MergeStatus "Source drive $sourceDrive is on disk number: $sourceDiskNumber"
                }
                if ($null -ne $targetDiskNumber) {
                    Add-MergeStatus "Target drive $targetDrive is on disk number: $targetDiskNumber"
                }

                # Kiểm tra nếu cả hai đều có thông tin và so sánh
                if ($null -ne $sourceDiskNumber -and $null -ne $targetDiskNumber) {
                    if ($sourceDiskNumber -ne $targetDiskNumber) {
                        Add-MergeStatus "Error: Source drive $sourceDrive (Disk $sourceDiskNumber) and target drive $targetDrive (Disk $targetDiskNumber) are not on the same physical disk."
                        Add-MergeStatus "You can only extend a volume with space from the same physical disk."
                        Add-MergeStatus "Operation aborted for safety."
                        return
                    }
                    Add-MergeStatus "Confirmed: Source and target drives are on the same physical disk (Disk $sourceDiskNumber). Proceeding..."
                } else {
                    # Nếu không thể xác định, hiển thị cảnh báo nhưng vẫn tiếp tục
                    Add-MergeStatus "Warning: Could not determine disk numbers for one or both drives. Will attempt to proceed anyway."
                    Add-MergeStatus "Note: This operation may fail if drives are not on the same physical disk."
                }
            }
            catch {
                Add-MergeStatus "Warning: Could not verify if drives are on the same physical disk. Proceeding anyway..."
                Add-MergeStatus "Error details: $_"
            }

            # Confirm operation
            $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                "WARNING: This will DELETE drive $sourceDrive and all its data, then extend drive $targetDrive.`n`nAre you sure you want to continue?",
                "Confirm Merge",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            if ($confirmResult -eq [System.Windows.Forms.DialogResult]::No) {
                Add-MergeStatus "Operation cancelled by user."
                return
            }

            # Create a batch file that will run diskpart
            $batchFilePath = "merge_volumes.bat"

            $batchContent = @"
@echo off
setlocal enabledelayedexpansion

REM Tạo file log
echo ============================================================ > merge_log.txt
echo                  Merging Volumes >> merge_log.txt
echo ============================================================ >> merge_log.txt
echo. >> merge_log.txt

REM Xóa ổ đĩa nguồn bằng PowerShell (nhanh hơn diskpart)
echo Deleting source drive %sourceDrive%... >> merge_log.txt

REM Sử dụng -WindowStyle Hidden để ẩn cửa sổ PowerShell
powershell -WindowStyle Hidden -command "& { try { Remove-Partition -DriveLetter %sourceDrive% -Confirm:$false -ErrorAction Stop; Write-Output 'Successfully deleted source drive %sourceDrive%.' } catch { Write-Error $_.Exception.Message; exit 1 } }" > delete_output.txt 2>&1

REM Chỉ ghi log, không hiển thị ra màn hình
type delete_output.txt >> merge_log.txt

REM Kiểm tra lỗi
if errorlevel 1 (
    echo PowerShell delete failed, trying diskpart... >> merge_log.txt

    REM Sử dụng diskpart nếu PowerShell thất bại
    (
        echo select volume %sourceDrive%
        echo delete volume override
    ) > diskpart_delete.txt

    REM Chạy diskpart ẩn
    start /b /wait "" cmd /c "diskpart /s diskpart_delete.txt > diskpart_delete_output.txt 2>&1"

    REM Chỉ ghi log, không hiển thị ra màn hình
    type diskpart_delete_output.txt >> merge_log.txt

    if errorlevel 1 (
        echo ERROR: Failed to delete source drive %sourceDrive%. >> merge_log.txt
        echo ERROR: Failed to delete source drive %sourceDrive%.
        echo This could be because the volume is in use or is a system volume. >> merge_log.txt
        echo This could be because the volume is in use or is a system volume.
        del diskpart_delete.txt
        del diskpart_delete_output.txt
        del delete_output.txt
        exit /b 1
    )

    del diskpart_delete.txt
    del diskpart_delete_output.txt
)

del delete_output.txt

REM Mở rộng ổ đĩa đích - sử dụng nhiều phương pháp
echo Extending target drive %targetDrive%... >> merge_log.txt

REM Đợi để hệ thống cập nhật
echo Waiting for system to update... >> merge_log.txt
timeout /t 2 /nobreak > nul

REM Phương pháp 1: Sử dụng diskpart (đáng tin cậy nhất)
echo Trying Method 1 (diskpart)... >> merge_log.txt

REM Lấy thông tin về disk number
echo Getting disk information... >> merge_log.txt
(
    echo list volume
) > diskpart_info.txt

REM Chạy diskpart ẩn
start /b /wait "" cmd /c "diskpart /s diskpart_info.txt > diskpart_info_output.txt 2>&1"
type diskpart_info_output.txt >> merge_log.txt

REM Tìm disk number của target drive
powershell -WindowStyle Hidden -command "& { $output = Get-Content -Path 'diskpart_info_output.txt'; $line = $output | Where-Object { $_ -match '%targetDrive%' }; if ($line -match 'Disk\s+(\d+)') { $matches[1] } else { 'Unknown' } }" > disk_number.txt
set /p DISK_NUMBER=<disk_number.txt

echo Target drive %targetDrive% is on disk %DISK_NUMBER% >> merge_log.txt

REM Tạo script diskpart để mở rộng ổ đĩa
(
    echo select disk %DISK_NUMBER%
    echo select volume %targetDrive%
    echo extend
) > diskpart_extend.txt

REM Chạy diskpart ẩn
start /b /wait "" cmd /c "diskpart /s diskpart_extend.txt > diskpart_extend_output.txt 2>&1"
type diskpart_extend_output.txt >> merge_log.txt

REM Kiểm tra lỗi
if errorlevel 1 (
    echo Method 1 failed, trying Method 2... >> merge_log.txt

    REM Phương pháp 2: Sử dụng PowerShell
    echo Trying Method 2 (PowerShell)... >> merge_log.txt

    REM Chạy PowerShell ẩn
    powershell -WindowStyle Hidden -command "& { try { $size = (Get-PartitionSupportedSize -DriveLetter %targetDrive%).SizeMax; Resize-Partition -DriveLetter %targetDrive% -Size $size -ErrorAction Stop; Write-Output 'Successfully extended partition using PowerShell.' } catch { Write-Error $_.Exception.Message; exit 1 } }" > extend_output.txt 2>&1

    REM Chỉ ghi log, không hiển thị ra màn hình
    type extend_output.txt >> merge_log.txt

    if errorlevel 1 (
        echo Method 2 failed, trying Method 3... >> merge_log.txt

        REM Phương pháp 3: Sử dụng diskpart với cách khác
        echo Trying Method 3 (alternative diskpart)... >> merge_log.txt

        REM Tạo script diskpart mới
        (
            echo rescan
            echo select volume %targetDrive%
            echo extend
        ) > diskpart_extend2.txt

        REM Chạy diskpart ẩn
        start /b /wait "" cmd /c "diskpart /s diskpart_extend2.txt > diskpart_extend2_output.txt 2>&1"

        REM Chỉ ghi log, không hiển thị ra màn hình
        type diskpart_extend2_output.txt >> merge_log.txt

        if errorlevel 1 (
            echo ERROR: Failed to extend target drive %targetDrive% using all methods. >> merge_log.txt
            echo This could be because there is no unallocated space adjacent to the volume. >> merge_log.txt
            del diskpart_info.txt
            del diskpart_info_output.txt
            del disk_number.txt
            del diskpart_extend.txt
            del diskpart_extend_output.txt
            del extend_output.txt
            del diskpart_extend2.txt
            del diskpart_extend2_output.txt
            exit /b 1
        )

        del diskpart_extend2.txt
        del diskpart_extend2_output.txt
        del extend_output.txt
    } else {
        echo Successfully extended partition using PowerShell. >> merge_log.txt
        del extend_output.txt
    }
) else (
    echo Successfully extended partition using diskpart. >> merge_log.txt
)

REM Xóa các file tạm
del diskpart_info.txt
del diskpart_info_output.txt
del disk_number.txt
del diskpart_extend.txt
del diskpart_extend_output.txt

del extend_output.txt

REM Hiển thị thông tin ổ đĩa sau khi hoàn thành
echo. >> merge_log.txt
echo Merge completed successfully! >> merge_log.txt
echo. >> merge_log.txt
echo ============================================================ >> merge_log.txt
echo                  Updated Drive List >> merge_log.txt
echo ============================================================ >> merge_log.txt
powershell -command "Get-CimInstance -ClassName Win32_LogicalDisk | Select-Object DeviceID, VolumeName, @{Name='Size (GB)';Expression={[math]::round(`$_.Size/1GB, 0)}}, @{Name='FreeSpace (GB)';Expression={[math]::round(`$_.FreeSpace/1GB, 0)}} | Format-Table -AutoSize" >> merge_log.txt
echo ============================================================ >> merge_log.txt

echo.
echo Merge completed successfully!
echo.
echo ============================================================
echo                  Updated Drive List
echo ============================================================
powershell -command "Get-CimInstance -ClassName Win32_LogicalDisk | Select-Object DeviceID, VolumeName, @{Name='Size (GB)';Expression={[math]::round(`$_.Size/1GB, 0)}}, @{Name='FreeSpace (GB)';Expression={[math]::round(`$_.FreeSpace/1GB, 0)}} | Format-Table -AutoSize"
echo ============================================================
echo.
echo Operation completed. You can close this window.
exit /b 0
"@
            Set-Content -Path $batchFilePath -Value $batchContent -Force -Encoding ASCII

            Add-MergeStatus "Merging volumes: deleting drive $sourceDrive and extending drive $targetDrive..."
            Add-MergeStatus "A command prompt window will open to complete the operation."
            Add-MergeStatus "Please follow the instructions in the command prompt window."

            try {
                # Hiển thị thông báo
                Add-MergeStatus "Merging volumes: deleting drive $sourceDrive and extending drive $targetDrive..."
                Add-MergeStatus "Please wait while the operation completes..."

                # Tạo và lưu batch file
                Set-Content -Path $batchFilePath -Value $batchContent -Force -Encoding ASCII

                # Hiển thị nội dung batch file để debug
                Add-MergeStatus "Batch file content:"
                Add-MergeStatus "-------------------"
                Add-MergeStatus (Get-Content -Path $batchFilePath -Raw)
                Add-MergeStatus "-------------------"

                # Chạy batch file với quyền admin và ẩn cửa sổ cmd
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = "cmd.exe"
                $psi.Arguments = "/c `"$batchFilePath`""
                $psi.UseShellExecute = $true
                $psi.Verb = "runas"
                $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

                # Chạy process
                try {
                    Add-MergeStatus "Starting batch process..."
                    $batchProcess = [System.Diagnostics.Process]::Start($psi)

                    # Hiển thị thông báo đang xử lý với thông tin chi tiết hơn
                    $progressCounter = 0
                    $progressChars = @('|', '/', '-', '\')
                    $progressSteps = @(
                        "Initializing operation...",
                        "Checking disk information...",
                        "Preparing to delete source drive...",
                        "Deleting source drive...",
                        "Waiting for system to update...",
                        "Preparing to extend target drive...",
                        "Extending target drive...",
                        "Finalizing operation..."
                    )
                    $currentStep = 0
                    $stepDuration = 0
                    $maxStepDuration = 8  # Số lần hiển thị mỗi bước trước khi chuyển sang bước tiếp theo

                    Add-MergeStatus "Starting volume merge operation..."

                    while (!$batchProcess.HasExited) {
                        # Cập nhật biểu tượng tiến trình
                        $progressChar = $progressChars[$progressCounter % 4]

                        # Cập nhật bước tiến trình
                        $stepDuration++
                        if ($stepDuration -ge $maxStepDuration) {
                            $currentStep = ($currentStep + 1) % $progressSteps.Count
                            $stepDuration = 0
                        }

                        # Hiển thị thông báo tiến trình
                        $currentMessage = $progressSteps[$currentStep]
                        Add-MergeStatus "$currentMessage $progressChar" -NoNewLine -ClearLine

                        # Cập nhật bộ đếm và đợi
                        $progressCounter++
                        [System.Windows.Forms.Application]::DoEvents()
                        Start-Sleep -Milliseconds 250
                    }

                    # Kiểm tra kết quả
                    if ($batchProcess.ExitCode -eq 0) {
                        Add-MergeStatus "Operation completed successfully!"
                    } else {
                        Add-MergeStatus "Batch process completed with exit code: $($batchProcess.ExitCode)"
                    }
                }
                catch {
                    Add-MergeStatus "Error starting batch process: $_"
                    return
                }

                # Đọc file log nếu có
                $logFilePath = "merge_log.txt"
                if (Test-Path $logFilePath) {
                    $logContent = Get-Content $logFilePath -Raw
                    Add-MergeStatus "---- Operation Log ----"
                    Add-MergeStatus $logContent
                    Add-MergeStatus "---- End of Log ----"
                    Remove-Item $logFilePath -Force -ErrorAction SilentlyContinue
                }

                # Check if successful
                if ($batchProcess.ExitCode -eq 0) {
                    Add-MergeStatus "Operation completed successfully."
                    Add-Status "Merged volumes: deleted drive $sourceDrive and extended drive $targetDrive."

                    # Refresh drive list
                    Add-MergeStatus "Refreshing drive list..."
                    Start-Sleep -Seconds 2

                    # Get updated list of drives
                    $updatedDrives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } },
                    @{Name = 'VolumeName'; Expression = { $_.VolumeName } },
                    @{Name = 'Size (GB)'; Expression = { [math]::round($_.Size / 1GB, 0) } },
                    @{Name = 'FreeSpace (GB)'; Expression = { [math]::round($_.FreeSpace / 1GB, 0) } }

                    # Display updated drive list
                    $driveListBox.Items.Clear()
                    foreach ($drive in $updatedDrives) {
                        $driveInfo = "$($drive.Name) - $($drive.VolumeName) - Size: $($drive.'Size (GB)') GB - Free: $($drive.'FreeSpace (GB)') GB"
                        $driveListBox.Items.Add($driveInfo)
                    }

                    if ($driveListBox.Items.Count -gt 0) {
                        $driveListBox.SelectedIndex = 0
                    }

                    # Clear the textboxes
                    $sourceDriveTextBox.Text = ""
                    $targetDriveTextBox.Text = ""
                }
                else {
                    # Thử thực hiện trực tiếp bằng PowerShell (phương pháp tối ưu)
                    Add-MergeStatus "Batch file failed with exit code $($batchProcess.ExitCode). Trying direct PowerShell method..."

                    # Tạo một hàm để xử lý việc mở rộng ổ đĩa
                    function Expand-Volume {
                        param (
                            [string]$SourceDrive,
                            [string]$TargetDrive
                        )

                        # Lấy thông tin về disk number
                        Add-MergeStatus "Getting disk information..."
                        $sourceDiskNumber = $null
                        $targetDiskNumber = $null

                        # Phương pháp 1: Sử dụng Get-Partition
                        try {
                            $sourcePartition = Get-Partition -DriveLetter $SourceDrive -ErrorAction Stop
                            $targetPartition = Get-Partition -DriveLetter $TargetDrive -ErrorAction Stop

                            $sourceDiskNumber = $sourcePartition.DiskNumber
                            $targetDiskNumber = $targetPartition.DiskNumber

                            Add-MergeStatus "Source drive $SourceDrive is on disk number: $sourceDiskNumber"
                            Add-MergeStatus "Target drive $TargetDrive is on disk number: $targetDiskNumber"
                        }
                        catch {
                            Add-MergeStatus "Error getting disk information: $_"
                        }

                        # Kiểm tra xem hai ổ đĩa có nằm trên cùng một đĩa vật lý không
                        if ($null -eq $sourceDiskNumber -or $null -eq $targetDiskNumber) {
                            Add-MergeStatus "Warning: Could not determine disk numbers for one or both drives. Will attempt to proceed anyway."
                        }
                        elseif ($sourceDiskNumber -ne $targetDiskNumber) {
                            Add-MergeStatus "Error: Source drive $SourceDrive (Disk $sourceDiskNumber) and target drive $TargetDrive (Disk $targetDiskNumber) are not on the same physical disk."
                            Add-MergeStatus "Volumes must be on the same physical disk to merge them."
                            Add-MergeStatus "Operation aborted for safety."
                            return $false
                        }
                        else {
                            Add-MergeStatus "Confirmed: Source and target drives are on the same physical disk (Disk $sourceDiskNumber)."
                        }

                        # Xóa ổ đĩa nguồn
                        Add-MergeStatus "Removing source drive $SourceDrive..."
                        try {
                            # Phương pháp 1: Sử dụng Remove-Partition
                            try {
                                # Lấy partition của ổ đĩa nguồn
                                $sourcePartition = Get-Partition -DriveLetter $SourceDrive -ErrorAction Stop
                                # Xóa partition
                                $sourcePartition | Remove-Partition -Confirm:$false -ErrorAction Stop
                                Add-MergeStatus "Source drive removed successfully using PowerShell."
                            }
                            catch {
                                Add-MergeStatus "PowerShell remove failed: $_"
                                Add-MergeStatus "Trying diskpart to remove source drive..."

                                # Phương pháp 2: Sử dụng diskpart
                                $diskpartScript = @"
select volume $SourceDrive
delete volume override
"@
                                Set-Content -Path "delete_volume.txt" -Value $diskpartScript

                                # Chạy diskpart ẩn và lấy output
                                $pInfo = New-Object System.Diagnostics.ProcessStartInfo
                                $pInfo.FileName = "diskpart.exe"
                                $pInfo.Arguments = "/s delete_volume.txt"
                                $pInfo.UseShellExecute = $false
                                $pInfo.RedirectStandardOutput = $true
                                $pInfo.CreateNoWindow = $true

                                $process = New-Object System.Diagnostics.Process
                                $process.StartInfo = $pInfo
                                $process.Start() | Out-Null
                                $diskpartOutput = $process.StandardOutput.ReadToEnd()
                                $process.WaitForExit()

                                Add-MergeStatus "Diskpart output: $diskpartOutput"

                                # Xóa file tạm
                                Remove-Item "delete_volume.txt" -ErrorAction SilentlyContinue

                                # Kiểm tra xem ổ đĩa đã bị xóa chưa
                                if (Get-Partition -DriveLetter $SourceDrive -ErrorAction SilentlyContinue) {
                                    Add-MergeStatus "Error: Failed to remove source drive using both methods."
                                    return $false
                                }

                                Add-MergeStatus "Source drive removed successfully using diskpart."
                            }
                        }
                        catch {
                            Add-MergeStatus "Error removing source drive: $_"
                            return $false
                        }

                        # Đợi một chút để hệ thống cập nhật
                        Add-MergeStatus "Waiting for system to update..."
                        Start-Sleep -Seconds 2

                        # Mở rộng ổ đĩa đích
                        Add-MergeStatus "Extending target drive $TargetDrive..."

                        # Phương pháp 1: Sử dụng Resize-Partition
                        try {
                            # Lấy kích thước tối đa có thể
                            $maxSize = (Get-PartitionSupportedSize -DriveLetter $TargetDrive).SizeMax

                            # Mở rộng partition
                            Resize-Partition -DriveLetter $TargetDrive -Size $maxSize -ErrorAction Stop
                            Add-MergeStatus "Target drive extended successfully using PowerShell."
                            return $true
                        }
                        catch {
                            Add-MergeStatus "PowerShell extend failed: $_"
                            Add-MergeStatus "Trying diskpart to extend target drive..."

                            # Phương pháp 2: Sử dụng diskpart
                            try {
                                $diskpartScript = @"
select volume $TargetDrive
extend
"@
                                Set-Content -Path "extend_volume.txt" -Value $diskpartScript

                                # Chạy diskpart ẩn và lấy output
                                $pInfo = New-Object System.Diagnostics.ProcessStartInfo
                                $pInfo.FileName = "diskpart.exe"
                                $pInfo.Arguments = "/s extend_volume.txt"
                                $pInfo.UseShellExecute = $false
                                $pInfo.RedirectStandardOutput = $true
                                $pInfo.CreateNoWindow = $true

                                $process = New-Object System.Diagnostics.Process
                                $process.StartInfo = $pInfo
                                $process.Start() | Out-Null
                                $diskpartOutput = $process.StandardOutput.ReadToEnd()
                                $process.WaitForExit()

                                Add-MergeStatus "Diskpart output: $diskpartOutput"

                                # Xóa file tạm
                                Remove-Item "extend_volume.txt" -ErrorAction SilentlyContinue

                                # Kiểm tra xem ổ đĩa đã được mở rộng chưa
                                Add-MergeStatus "Target drive extended successfully using diskpart."
                                return $true
                            }
                            catch {
                                Add-MergeStatus "Diskpart extend failed: $_"
                                Add-MergeStatus "Error extending target drive using both methods."
                                return $false
                            }
                        }
                    }

                    # Gọi hàm mở rộng ổ đĩa
                    $success = Expand-Volume -SourceDrive $sourceDrive -TargetDrive $targetDrive

                    if ($success) {
                        Add-Status "Merged volumes: deleted drive $sourceDrive and extended drive $targetDrive."

                        # Refresh drive list
                        Add-MergeStatus "Refreshing drive list..."
                        Start-Sleep -Seconds 1

                        # Sử dụng Get-CimInstance thay vì Get-WmiObject để cải thiện hiệu suất
                        $updatedDrives = Get-CimInstance -ClassName Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } },
                        @{Name = 'VolumeName'; Expression = { $_.VolumeName } },
                        @{Name = 'Size (GB)'; Expression = { [math]::round($_.Size / 1GB, 0) } },
                        @{Name = 'FreeSpace (GB)'; Expression = { [math]::round($_.FreeSpace / 1GB, 0) } }

                        # Cập nhật danh sách ổ đĩa
                        $driveListBox.Items.Clear()
                        foreach ($drive in $updatedDrives) {
                            $driveInfo = "$($drive.Name) - $($drive.VolumeName) - Size: $($drive.'Size (GB)') GB - Free: $($drive.'FreeSpace (GB)') GB"
                            $driveListBox.Items.Add($driveInfo)
                        }

                        if ($driveListBox.Items.Count -gt 0) {
                            $driveListBox.SelectedIndex = 0
                        }

                        # Xóa nội dung textbox
                        $sourceDriveTextBox.Text = ""
                        $targetDriveTextBox.Text = ""
                    }
                    else {
                        Add-MergeStatus "The operation failed. Please try using Disk Management instead."
                    }
                }
            }
            catch {
                Add-MergeStatus "Error: $_"
            }
            finally {
                # Clean up temp file
                if (Test-Path $batchFilePath) {
                    Remove-Item $batchFilePath -Force -ErrorAction SilentlyContinue
                }
            }
        }
        $contentPanel.Controls.Add($mergeButton)

        # Cancel button
        $cancelButton = New-DynamicButton -text "Cancel" -x 410 -y 140 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
            $contentPanel.Controls.Clear()
        }
        $contentPanel.Controls.Add($cancelButton)

        Add-Status "Ready to extend volume. Select source and target drives, then click Merge and Extend Volume."

    }
    $volumeForm.Controls.Add($btnExtendVolume)

    # Rename Volume button
    $btnRenameVolume = New-DynamicButton -text "Rename Volume" -x 500 -y 150 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Clear the content panel
        $contentPanel.Controls.Clear()

        # Title label
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "Rename Volume"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 10)
        $titleLabel.Size = New-Object System.Drawing.Size(760, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $contentPanel.Controls.Add($titleLabel)

        # Note: We're using the driveListBox that's already in the parent form

        # Drive letter label
        $driveLetterLabel = New-Object System.Windows.Forms.Label
        $driveLetterLabel.Text = "Drive Letter:"
        $driveLetterLabel.Location = New-Object System.Drawing.Point(200, 50)
        $driveLetterLabel.Size = New-Object System.Drawing.Size(100, 20)
        $driveLetterLabel.ForeColor = [System.Drawing.Color]::White
        $driveLetterLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $contentPanel.Controls.Add($driveLetterLabel)

        # Drive letter textbox
        $driveLetterTextBox = New-Object System.Windows.Forms.TextBox
        $driveLetterTextBox.Location = New-Object System.Drawing.Point(310, 50)
        $driveLetterTextBox.Size = New-Object System.Drawing.Size(50, 20)
        $driveLetterTextBox.BackColor = [System.Drawing.Color]::White
        $driveLetterTextBox.ForeColor = [System.Drawing.Color]::Black
        $driveLetterTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $driveLetterTextBox.MaxLength = 1
        $driveLetterTextBox.ReadOnly = $true
        $driveLetterTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
        $contentPanel.Controls.Add($driveLetterTextBox)

        # Update drive letter from selected drive
        if ($driveListBox.SelectedItem) {
            $selectedDrive = $driveListBox.SelectedItem.ToString()
            $driveLetter = $selectedDrive.Substring(0, 1)
            $driveLetterTextBox.Text = $driveLetter

            # Get current volume name
            $drives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } },
            @{Name = 'VolumeName'; Expression = { $_.VolumeName } }

            $currentVolumeName = ""
            foreach ($drive in $drives) {
                if ($drive.Name -eq "$($driveLetter):") {
                    $currentVolumeName = $drive.VolumeName
                    break
                }
            }
        }

        # New label label
        $newLabelLabel = New-Object System.Windows.Forms.Label
        $newLabelLabel.Text = "New Label:"
        $newLabelLabel.Location = New-Object System.Drawing.Point(200, 80)
        $newLabelLabel.Size = New-Object System.Drawing.Size(100, 20)
        $newLabelLabel.ForeColor = [System.Drawing.Color]::White
        $newLabelLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $contentPanel.Controls.Add($newLabelLabel)

        # New label textbox
        $newLabelTextBox = New-Object System.Windows.Forms.TextBox
        $newLabelTextBox.Location = New-Object System.Drawing.Point(310, 80)
        $newLabelTextBox.Size = New-Object System.Drawing.Size(250, 20)
        $newLabelTextBox.BackColor = [System.Drawing.Color]::White
        $newLabelTextBox.ForeColor = [System.Drawing.Color]::Black
        $newLabelTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $newLabelTextBox.Text = $currentVolumeName

        # Add event handler for Enter key
        $newLabelTextBox.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true  # Prevent beep sound
                $renameButton.PerformClick()  # Activate Rename button
            }
        })

        $contentPanel.Controls.Add($newLabelTextBox)

        # Function to add status message to the main status box for rename operations
        function Add-RenameStatus {
            param([string]$message)
            # Use the main status box
            Add-Status $message
        }

        # Note: The driveListBox.SelectedIndexChanged event is now handled at the form level

        # Hàm để lấy thông tin ổ đĩa
        function Get-DriveInfo {
            return Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } },
            @{Name = 'VolumeName'; Expression = { $_.VolumeName } },
            @{Name = 'Size (GB)'; Expression = { [math]::round($_.Size / 1GB, 0) } },
            @{Name = 'FreeSpace (GB)'; Expression = { [math]::round($_.FreeSpace / 1GB, 0) } }
        }

        # Hàm để cập nhật danh sách ổ đĩa trong giao diện
        function Update-DriveList {
            param (
                [Parameter(Mandatory = $true)]
                [array]$Drives,

                [Parameter(Mandatory = $false)]
                [string]$SelectDriveLetter = ""
            )

            # Xóa danh sách hiện tại
            $driveListBox.Items.Clear()

            # Thêm các ổ đĩa vào danh sách
            foreach ($drive in $Drives) {
                $driveInfo = "$($drive.Name) - $($drive.VolumeName) - Size: $($drive.'Size (GB)') GB - Free: $($drive.'FreeSpace (GB)') GB"
                $driveListBox.Items.Add($driveInfo)
            }

            # Chọn ổ đĩa trong danh sách
            if ($SelectDriveLetter -ne "") {
                # Tìm ổ đĩa cần chọn
                $selectedIndex = -1
                for ($i = 0; $i -lt $driveListBox.Items.Count; $i++) {
                    $item = $driveListBox.Items[$i].ToString()
                    if ($item.StartsWith("$($SelectDriveLetter):")) {
                        $selectedIndex = $i
                        break
                    }
                }

                # Chọn ổ đĩa
                if ($selectedIndex -ge 0) {
                    $driveListBox.SelectedIndex = $selectedIndex
                } elseif ($driveListBox.Items.Count -gt 0) {
                    $driveListBox.SelectedIndex = 0
                }
            } elseif ($driveListBox.Items.Count -gt 0) {
                $driveListBox.SelectedIndex = 0
            }

            # Cập nhật biến toàn cục
            $script:drives = $Drives
        }

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

        # Hàm tạo batch file để đổi tên ổ đĩa
        function New-RenameBatchFile {
            param (
                [Parameter(Mandatory = $true)]
                [string]$DriveLetter,

                [Parameter(Mandatory = $true)]
                [string]$NewLabel
            )

            # Đảm bảo không có khoảng trắng ở đầu và cuối tên nhãn
            $NewLabel = $NewLabel.Trim()

            $batchFilePath = [System.IO.Path]::GetTempFileName() + ".bat"

            $batchContent = @"
                @echo off
                echo ============================================================
                echo                  Renaming Volume $DriveLetter
                echo ============================================================
                echo.

                echo Checking if drive exists...
                powershell -command "(Get-WmiObject Win32_LogicalDisk).DeviceID -contains '$DriveLetter`:'" >nul 2>&1
                if errorlevel 1 (
                    echo ERROR: Drive $DriveLetter`: does not exist. Exiting...
                    pause
                    exit /b 1
                )

                echo Setting label for drive $DriveLetter`: to $NewLabel...
                @REM Sử dụng PowerShell trực tiếp để đổi tên ổ đĩa
                powershell -Command "& {
                    `$volume = Get-WmiObject -Query \"SELECT * FROM Win32_Volume WHERE DriveLetter='$DriveLetter`:'\"
                    if (`$volume) {
                        `$volume.Label = '$NewLabel'
                        `$result = `$volume.Put()
                        if (`$result.ReturnValue -eq 0) { exit 0 } else { exit 1 }
                    } else {
                        `$logicalDisk = Get-WmiObject -Query \"SELECT * FROM Win32_LogicalDisk WHERE DeviceID='$DriveLetter`:'\"
                        if (`$logicalDisk) {
                            `$logicalDisk.VolumeName = '$NewLabel'
                            `$result = `$logicalDisk.Put()
                            if (`$result.ReturnValue -eq 0) { exit 0 } else { exit 1 }
                        } else {
                            exit 1
                        }
                    }
                }"
                if errorlevel 1 (
                    echo ERROR: Failed to rename the drive. Check the drive letter and try again.
                    pause
                    exit /b 1
                ) else (
                    echo Drive $DriveLetter`: successfully renamed to $NewLabel.
                )

                echo.
                echo ============================================================
                echo                  Updated Drive List
                echo ============================================================
                powershell -command "Get-WmiObject Win32_LogicalDisk | Select-Object @{Name='Name';Expression={`$_.DeviceID}}, @{Name='VolumeName';Expression={`$_.VolumeName}}, @{Name='Size (GB)';Expression={[math]::round(`$_.Size/1GB, 0)}}, @{Name='FreeSpace (GB)';Expression={[math]::round(`$_.FreeSpace/1GB, 0)}} | Format-Table -AutoSize"
                echo ============================================================
                echo.
                pause
"@
            Set-Content -Path $batchFilePath -Value $batchContent -Force -Encoding ASCII

            return $batchFilePath
        }

        # Rename button
        $renameButton = New-DynamicButton -text "Rename Volume" -x 150 -y 140 -width 200 -height 30 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            # Lấy thông tin từ form
            $driveLetter = $driveLetterTextBox.Text.Trim().ToUpper()
            # Đảm bảo không có khoảng trắng ở đầu và cuối tên nhãn
            $newLabel = $newLabelTextBox.Text.Trim()

            # Kiểm tra đầu vào
            if ($driveLetter -eq "") {
                Add-Status "Rename Volume: Error: Please select a drive."
                return
            }

            if ($newLabel -eq "") {
                Add-Status "Rename Volume: Error: Please enter a new label."
                return
            }

            # Thông báo bắt đầu đổi tên
            Add-Status "Rename Volume: Renaming drive $driveLetter to $newLabel..."

            # Biến để lưu trạng thái thành công
            $success = $false
            $batchFilePath = $null

            try {
                # Thử đổi tên trực tiếp trước
                Add-RenameStatus "Attempting to rename directly..."

                # Thử đổi tên bằng WMI
                if (Rename-DriveWithWMI -DriveLetter $driveLetter -NewLabel $newLabel) {
                    Add-RenameStatus "Successfully renamed drive $driveLetter to $newLabel using WMI."
                    $success = $true
                }
                # Thử đổi tên bằng lệnh label
                elseif (Rename-DriveWithLabel -DriveLetter $driveLetter -NewLabel $newLabel) {
                    Add-RenameStatus "Successfully renamed drive $driveLetter to $newLabel using label command."
                    $success = $true
                }
                # Nếu không thành công, tạo và chạy batch file
                else {
                    Add-RenameStatus "Direct rename failed. Trying with batch file..."

                    # Tạo batch file
                    $batchFilePath = New-RenameBatchFile -DriveLetter $driveLetter -NewLabel $newLabel

                    # Thông báo
                    Add-RenameStatus "A command prompt window will open to complete the operation."
                    Add-RenameStatus "Please follow the instructions in the command prompt window."

                    # Chạy batch file với quyền admin
                    $batchProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$batchFilePath`"" -Verb RunAs -PassThru -Wait

                    # Kiểm tra kết quả
                    if ($batchProcess.ExitCode -eq 0) {
                        Add-RenameStatus "Operation completed successfully."
                        Add-Status "Renamed drive $driveLetter to $newLabel."
                        $success = $true
                    }
                    else {
                        Add-RenameStatus "Error: The operation failed with exit code $($batchProcess.ExitCode)"
                        Add-RenameStatus "Please check the command prompt window for more details."
                    }
                }

                # Nếu thành công, cập nhật danh sách ổ đĩa
                if ($success) {
                    # Cập nhật danh sách ổ đĩa
                    Add-RenameStatus "Refreshing drive list..."
                    Start-Sleep -Seconds 1

                    # Thử lấy thông tin ổ đĩa nhiều lần để đảm bảo cập nhật
                    $maxRetries = 3
                    $retryCount = 0
                    $updatedDrives = $null
                    $driveUpdated = $false

                    while ($retryCount -lt $maxRetries -and -not $driveUpdated) {
                        # Lấy danh sách ổ đĩa mới
                        $updatedDrives = Get-DriveInfo

                        # Kiểm tra xem tên ổ đĩa đã được cập nhật chưa
                        foreach ($drive in $updatedDrives) {
                            # Kiểm tra chính xác tên ổ đĩa và tên nhãn
                            if ($drive.Name -eq "$($driveLetter):" -and $drive.VolumeName -eq $newLabel) {
                                $driveUpdated = $true
                                break
                            }
                        }

                        if (-not $driveUpdated) {
                            Add-RenameStatus "Waiting for drive information to update..."
                            Start-Sleep -Seconds 1
                            $retryCount++
                        }
                    }

                    # Cập nhật giao diện
                    Update-DriveList -Drives $updatedDrives -SelectDriveLetter $driveLetter
                    Add-RenameStatus "Drive list updated successfully."
                }
            }
            catch {
                Add-RenameStatus "Error: $_"
            }
            finally {
                # Xóa file tạm nếu có
                if ($batchFilePath -and (Test-Path $batchFilePath)) {
                    Remove-Item $batchFilePath -Force -ErrorAction SilentlyContinue
                }
            }
        }
        $contentPanel.Controls.Add($renameButton)

        # Cancel button
        $cancelButton = New-DynamicButton -text "Cancel" -x 410 -y 140 -width 200 -height 30 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
            $contentPanel.Controls.Clear()
        }
        $contentPanel.Controls.Add($cancelButton)

        Add-Status "Ready to rename volume. Select a drive from the list, enter a new label, then click Rename Volume."
    }
    $volumeForm.Controls.Add($btnRenameVolume)

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

# Functions to hide and show the main menu
function Hide-MainMenu {
    $script:form.Hide()
}

# Function to show the main menu
function Show-MainMenu {
    $script:form.Show()
    $script:form.BringToFront()
}

# Add buttons to form
$form.Controls.Add($buttonRunAll)
$form.Controls.Add($buttonInstallSoftware)
$form.Controls.Add($buttonPowerOptions)
$form.Controls.Add($buttonChangeVolume)
$form.Controls.Add($buttonActivate)
$form.Controls.Add($buttonTurnOnFeatures)
$form.Controls.Add($buttonRenameDevice)
$form.Controls.Add($buttonSetPassword)
$form.Controls.Add($buttonJoinDomain)
$form.Controls.Add($buttonExit)

# Display form
$form.ShowDialog()