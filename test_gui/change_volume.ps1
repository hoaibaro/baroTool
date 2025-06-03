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

                if ($driveLetterTextBox) {
                    $driveLetterTextBox.Text = $driveLetter
                    # Don't auto-fill the new label textbox - let user type what they want
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

                    # Find newly created drive
                    $currentDrives = Get-WmiObject Win32_LogicalDisk | Select-Object DeviceID, VolumeName
                    foreach ($drive in $currentDrives) {
                        if ($drive.DeviceID -ne "$($driveLetter):" -and
                            ($drive.VolumeName -eq "New Volume" -or $drive.VolumeName -eq "")) {
                            $newDriveFound = $true
                            $newDriveLetter = $drive.DeviceID.TrimEnd(":")
                            Add-Status "Found new drive: $newDriveLetter"
                            break
                        }
                    }

                    # Rename the new drive if found
                    if ($newDriveFound) {
                        $actualNewLabel = if (-not [string]::IsNullOrEmpty($newLabel)) { $newLabel } else { "GAME" }
                        Add-Status "Renaming new drive $newDriveLetter to $actualNewLabel..."

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
                    Add-Status "Operation completed. Found $driveCount drives."
                }
                else {
                    Add-Status "Operation completed with warnings. Check the log above."
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

        # Merge button
        $mergeButton = New-DynamicButton -text "Merge and Extend Volume" -x 150 -y 140 -width 250 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            # Input validation
            $sourceDrive = $sourceDriveTextBox.Text.Trim().ToUpper()
            $targetDrive = $targetDriveTextBox.Text.Trim().ToUpper()

            # Basic validation
            if ([string]::IsNullOrEmpty($sourceDrive)) {
                Add-Status "Error: Please enter a source drive letter."
                return
            }
            if ([string]::IsNullOrEmpty($targetDrive)) {
                Add-Status "Error: Please enter a target drive letter."
                return
            }
            if (-not ($sourceDrive -match '^[A-Z]$') -or -not ($targetDrive -match '^[A-Z]$')) {
                Add-Status "Error: Drive letters must be single letters (A-Z)."
                return
            }
            if ($sourceDrive -eq $targetDrive) {
                Add-Status "Error: Source and target drives cannot be the same."
                return
            }

            # Check if drives exist
            $existingDrives = Get-CimInstance -ClassName Win32_LogicalDisk | Select-Object -ExpandProperty DeviceID | ForEach-Object { $_.Substring(0, 1) }
            if ($existingDrives -notcontains $sourceDrive) {
                Add-Status "Error: Source drive $sourceDrive does not exist."
                return
            }
            if ($existingDrives -notcontains $targetDrive) {
                Add-Status "Error: Target drive $targetDrive does not exist."
                return
            }

            # Check if drives are on same physical disk (simplified)
            try {
                $sourcePartition = Get-Partition -DriveLetter $sourceDrive -ErrorAction Stop
                $targetPartition = Get-Partition -DriveLetter $targetDrive -ErrorAction Stop

                if ($sourcePartition.DiskNumber -ne $targetPartition.DiskNumber) {
                    Add-Status "Error: Drives are not on the same physical disk. Operation aborted for safety."
                    return
                }
                Add-Status "Drives are on the same physical disk. Proceeding..."
            }
            catch {
                Add-Status "Warning: Could not verify disk compatibility. Proceeding anyway..."
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

            # Perform merge operation using PowerShell
            Add-Status "Merging volumes: deleting drive $sourceDrive and extending drive $targetDrive..."

            try {
                # Delete source drive
                Add-Status "Removing source drive $sourceDrive..."
                Remove-Partition -DriveLetter $sourceDrive -Confirm:$false -ErrorAction Stop
                Add-Status "Source drive removed successfully."

                # Wait for system to update
                Start-Sleep -Seconds 2

                # Extend target drive
                Add-Status "Extending target drive $targetDrive..."
                $maxSize = (Get-PartitionSupportedSize -DriveLetter $targetDrive).SizeMax
                Resize-Partition -DriveLetter $targetDrive -Size $maxSize -ErrorAction Stop
                Add-Status "Target drive extended successfully."

                # Update drive list
                $driveCount = Update-DriveList
                Add-Status "Operation completed successfully. Found $driveCount drives."

                # Clear textboxes
                $sourceDriveTextBox.Text = ""
                $targetDriveTextBox.Text = ""
            }
            catch {
                Add-Status "Error: $_"
                Add-Status "Operation failed. Please try using Disk Management instead."
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
        }

        # New label label
        $newLabelLabel = New-Object System.Windows.Forms.Label
        $newLabelLabel.Text = "New Label:"
        $newLabelLabel.Location = New-Object System.Drawing.Point(200, 80)
        $newLabelLabel.Size = New-Object System.Drawing.Size(100, 20)
        $newLabelLabel.ForeColor = [System.Drawing.Color]::White
        $newLabelLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $contentPanel.Controls.Add($newLabelLabel)

        # New label textbox - use script scope
        $script:newLabelTextBox = New-Object System.Windows.Forms.TextBox
        $script:newLabelTextBox.Location = New-Object System.Drawing.Point(310, 80)
        $script:newLabelTextBox.Size = New-Object System.Drawing.Size(250, 20)
        $script:newLabelTextBox.BackColor = [System.Drawing.Color]::White
        $script:newLabelTextBox.ForeColor = [System.Drawing.Color]::Black
        $script:newLabelTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        # Let user type the new label they want
        $script:newLabelTextBox.Text = ""
        $script:newLabelTextBox.PlaceholderText = "Enter new volume label"

        # Add event handler for Enter key
        $script:newLabelTextBox.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true  # Prevent beep sound
                $renameButton.PerformClick()  # Activate Rename button
            }
        })

        $contentPanel.Controls.Add($script:newLabelTextBox)

        # Rename button
        $renameButton = New-DynamicButton -text "Rename Volume" -x 150 -y 140 -width 200 -height 30 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            # Get drive letter from selected item directly
            if (-not $driveListBox.SelectedItem) {
                Add-Status "Error: Please select a drive from the list first."
                return
            }

            $selectedDrive = $driveListBox.SelectedItem.ToString()
            $driveLetter = $selectedDrive.Substring(0, 1).ToUpper()
            $newLabel = $script:newLabelTextBox.Text.Trim()

            if ($newLabel -eq "") {
                Add-Status "Error: Please enter a new label."
                return
            }

            # Rename volume using multiple methods
            Add-Status "Renaming drive $driveLetter to $newLabel..."

            # Try Set-Volume first (Windows 8+)
            try {
                if (Get-Command Set-Volume -ErrorAction SilentlyContinue) {
                    Set-Volume -DriveLetter $driveLetter -NewFileSystemLabel $newLabel -ErrorAction Stop
                    Add-Status "Successfully renamed drive $driveLetter to $newLabel using Set-Volume."
                    
                    # Update drive list
                    $driveCount = Update-DriveList
                    Add-Status "Drive list updated. Found $driveCount drives."
                    
                    # Clear textbox
                    $script:newLabelTextBox.Text = ""
                    return
                }
            }
            catch {
                Add-Status "Set-Volume failed, trying alternative methods..."
            }

            # Try WMI method
            if (Rename-DriveWithWMI -DriveLetter $driveLetter -NewLabel $newLabel) {
                Add-Status "Successfully renamed drive $driveLetter to $newLabel using WMI."
                
                # Update drive list
                $driveCount = Update-DriveList
                Add-Status "Drive list updated. Found $driveCount drives."
                
                # Clear textbox
                $script:newLabelTextBox.Text = ""
                return
            }

            # Try label command as last resort
            if (Rename-DriveWithLabel -DriveLetter $driveLetter -NewLabel $newLabel) {
                Add-Status "Successfully renamed drive $driveLetter to $newLabel using label command."
                
                # Update drive list
                $driveCount = Update-DriveList
                Add-Status "Drive list updated. Found $driveCount drives."
                
                # Clear textbox
                $script:newLabelTextBox.Text = ""
                return
            }

            # If all methods fail
            Add-Status "Failed to rename drive $driveLetter. Please try renaming manually to '$newLabel'."
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