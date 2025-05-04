# Check if running as administrator and restart with elevation if not
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
        [System.Drawing.Color]::FromArgb(0, 0, 0),  # Black at top
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
    } else {
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

# Run All Options
$buttonRunAll = New-DynamicButton -text "Run All Options" -x 30 -y 100 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    [System.Windows.Forms.MessageBox]::Show("Running all options...")
    # Add commands to run here
}

# Install All Software
$buttonInstallSoftware = New-DynamicButton -text "Install All Software" -x 30 -y 180 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    [System.Windows.Forms.MessageBox]::Show("Installing all software...")
    # Add installation commands here
}

# Power Options and Firewall
$buttonPowerOptions = New-DynamicButton -text "Power Options and Firewall" -x 30 -y 260 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Create Power Options and Firewall form
    $powerForm = New-Object System.Windows.Forms.Form
    $powerForm.Text = "Power Options and Firewall"
    $powerForm.Size = New-Object System.Drawing.Size(500, 550)
    $powerForm.StartPosition = "CenterScreen"
    $powerForm.BackColor = [System.Drawing.Color]::Black
    $powerForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $powerForm.MaximizeBox = $false
    $powerForm.MinimizeBox = $false

    # Add a gradient background
    $powerForm.Paint = {
        $graphics = $_.Graphics
        $rect = New-Object System.Drawing.Rectangle(0, 0, $powerForm.Width, $powerForm.Height)
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $rect,
            [System.Drawing.Color]::FromArgb(0, 0, 0),  # Black at top
            [System.Drawing.Color]::FromArgb(0, 40, 0), # Dark green at bottom
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
        )
        $graphics.FillRectangle($brush, $rect)
        $brush.Dispose()
    }

    # Title label with animation
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "POWER OPTIONS AND FIREWALL MANAGEMENT"
    $titleLabel.Location = New-Object System.Drawing.Point(-10, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(500, 40)
    $titleLabel.ForeColor = [System.Drawing.Color]::Lime
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $titleLabel.Padding = New-Object System.Windows.Forms.Padding(5)

    # Add animation to the title
    $titleTimer = New-Object System.Windows.Forms.Timer
    $titleTimer.Interval = 500
    $titleTimer.Add_Tick({
        if ($titleLabel.ForeColor -eq [System.Drawing.Color]::Lime) {
            $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 220, 0)
        } else {
            $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        }
    })
    $titleTimer.Start()

    $powerForm.Controls.Add($titleLabel)

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(50, 400)
    $statusTextBox.Size = New-Object System.Drawing.Size(400, 100)
    $statusTextBox.BackColor = [System.Drawing.Color]::Black
    $statusTextBox.ForeColor = [System.Drawing.Color]::Lime
    $statusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $statusTextBox.ReadOnly = $true

    # Add a border to the status text box
    $statusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    # Add a placeholder text
    $statusTextBox.Text = "Status messages will appear here..."

    $powerForm.Controls.Add($statusTextBox)

    # Function to add status message
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

    # Set Time/Timezone and Power Options button
    $btnTimeAndPower = New-DynamicButton -text "Set Time/Timezone and Power Options" -x 50 -y 80 -width 400 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        try {
            Add-Status "Setting time zone to SE Asia Standard Time..."

            # Set timezone to SE Asia Standard Time
            Start-Process -FilePath "tzutil.exe" -ArgumentList "/s `"SE Asia Standard Time`"" -Wait -NoNewWindow

            # Configure Windows Time service
            Add-Status "Configuring Windows Time service..."
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\w32time\Parameters" -Name "Type" -Value "NTP" -Type String -ErrorAction SilentlyContinue

            # Resync time
            Add-Status "Synchronizing time..."
            try {
                Start-Process -FilePath "w32tm.exe" -ArgumentList "/resync" -Wait -NoNewWindow
            }
            catch {
                Add-Status "Warning: Could not sync time. $_"
            }

            # Enable automatic time zone updates
            Add-Status "Enabling automatic time zone updates..."
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\tzautoupdate" -Name "Start" -Value 2 -Type DWord -ErrorAction SilentlyContinue

            # Configure power options
            Add-Status "Setting power options to 'Do Nothing'..."

            # Create a process to run the power commands with elevated privileges
            $powerCommands = @(
                "powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_BUTTONS LIDACTION 0",
                "powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_BUTTONS LIDACTION 0",
                "powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_BUTTONS SBUTTONACTION 0",
                "powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_BUTTONS SBUTTONACTION 0",
                "powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_BUTTONS PBUTTONACTION 0",
                "powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_BUTTONS PBUTTONACTION 0",
                "powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 0",
                "powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 0",
                "powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_SLEEP STANDBYIDLE 0",
                "powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_SLEEP STANDBYIDLE 0",
                "powercfg /SETACTIVE SCHEME_CURRENT"
            )

            $powerScript = $powerCommands -join "; "

            # Create a process to run the command with elevated privileges
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-Command Start-Process cmd.exe -ArgumentList '/c $powerScript' -Verb RunAs -WindowStyle Hidden"
            $psi.UseShellExecute = $true
            $psi.Verb = "runas"
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

            # Start the process
            [System.Diagnostics.Process]::Start($psi)

            # Turn off the firewall
            Add-Status "Turning off the firewall..."
            $command = "netsh advfirewall set allprofiles state off"

            # Create a process to run the command with elevated privileges
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-Command Start-Process cmd.exe -ArgumentList '/c $command' -Verb RunAs -WindowStyle Hidden"
            $psi.UseShellExecute = $true
            $psi.Verb = "runas"
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

            # Start the process
            [System.Diagnostics.Process]::Start($psi)

            Add-Status "Time zone, power options, and firewall have been configured successfully!"
        }
        catch {
            Add-Status "Error: $_"
        }
    }
    $powerForm.Controls.Add($btnTimeAndPower)

    # Turn on Firewall button
    $btnFirewallOn = New-DynamicButton -text "Turn on Firewall" -x 50 -y 160 -width 400 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        try {
            Add-Status "Turning on the firewall..."
            # Turn on the firewall
            $command = "netsh advfirewall set allprofiles state on"

            # Create a process to run the command with elevated privileges
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-Command Start-Process cmd.exe -ArgumentList '/c $command' -Verb RunAs -WindowStyle Hidden"
            $psi.UseShellExecute = $true
            $psi.Verb = "runas"
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

            # Start the process
            [System.Diagnostics.Process]::Start($psi)

            Add-Status "Firewall has been turned on successfully!"
        }
        catch {
            Add-Status "Error: $_"
        }
    }
    $powerForm.Controls.Add($btnFirewallOn)

    # Turn off Firewall button
    $btnFirewallOff = New-DynamicButton -text "Turn off Firewall" -x 50 -y 240 -width 400 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        try {
            Add-Status "Turning off the firewall..."
            # Turn off the firewall
            $command = "netsh advfirewall set allprofiles state off"

            # Create a process to run the command with elevated privileges
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-Command Start-Process cmd.exe -ArgumentList '/c $command' -Verb RunAs -WindowStyle Hidden"
            $psi.UseShellExecute = $true
            $psi.Verb = "runas"
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

            # Start the process
            [System.Diagnostics.Process]::Start($psi)

            Add-Status "Firewall has been turned off successfully!"
        }
        catch {
            Add-Status "Error: $_"
        }
    }
    $powerForm.Controls.Add($btnFirewallOff)

    # Return to Main Menu button
    $btnReturn = New-DynamicButton -text "Return to Main Menu" -x 50 -y 320 -width 400 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        $powerForm.Close()
    }
    $powerForm.Controls.Add($btnReturn)

    # Show the form
    $powerForm.ShowDialog()
}

# Change / Edit Volume
$buttonChangeVolume = New-DynamicButton -text "Change / Edit Volume" -x 30 -y 340 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Create volume management form
    $volumeForm = New-Object System.Windows.Forms.Form
    $volumeForm.Text = "Volume Management"
    $volumeForm.Size = New-Object System.Drawing.Size(500, 500)
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
            [System.Drawing.Color]::FromArgb(0, 0, 0),  # Black at top
            [System.Drawing.Color]::FromArgb(0, 40, 0), # Dark green at bottom
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
        )
        $graphics.FillRectangle($brush, $rect)
        $brush.Dispose()
    }

    # Title label with animation
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "CHANGE THE DRIVE LETTER"
    $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(500, 40)
    $titleLabel.ForeColor = [System.Drawing.Color]::Lime
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $titleLabel.Padding = New-Object System.Windows.Forms.Padding(5)

    # Add animation to the title
    $titleTimer = New-Object System.Windows.Forms.Timer
    $titleTimer.Interval = 500
    $titleTimer.Add_Tick({
        if ($titleLabel.ForeColor -eq [System.Drawing.Color]::Lime) {
            $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 220, 0)
        } else {
            $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        }
    })
    $titleTimer.Start()

    $volumeForm.Controls.Add($titleLabel)

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(50, 330)
    $statusTextBox.Size = New-Object System.Drawing.Size(400, 120)
    $statusTextBox.BackColor = [System.Drawing.Color]::Black
    $statusTextBox.ForeColor = [System.Drawing.Color]::Lime
    $statusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $statusTextBox.ReadOnly = $true
    $volumeForm.Controls.Add($statusTextBox)

    # Status text box with improved styling
    $statusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $statusTextBox.Text = "Status messages will appear here..."

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

    # Change Drive Letter button
    $btnChangeDriveLetter = New-DynamicButton -text "[1] Change the drive letter" -x 50 -y 80 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Create Change Drive Letter form
        $changeDriveForm = New-Object System.Windows.Forms.Form
        $changeDriveForm.Text = "Change Drive Letter"
        $changeDriveForm.Size = New-Object System.Drawing.Size(600, 500)
        $changeDriveForm.StartPosition = "CenterScreen"
        $changeDriveForm.BackColor = [System.Drawing.Color]::Black
        $changeDriveForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $changeDriveForm.MaximizeBox = $false
        $changeDriveForm.MinimizeBox = $false

        # Title label
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "Change Drive Letter"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(600, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $changeDriveForm.Controls.Add($titleLabel)

        # Drive list label
        $driveListLabel = New-Object System.Windows.Forms.Label
        $driveListLabel.Text = "Available Drives:"
        $driveListLabel.Location = New-Object System.Drawing.Point(20, 60)
        $driveListLabel.Size = New-Object System.Drawing.Size(200, 20)
        $driveListLabel.ForeColor = [System.Drawing.Color]::White
        $driveListLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $changeDriveForm.Controls.Add($driveListLabel)

        # Drive list box
        $driveListBox = New-Object System.Windows.Forms.ListBox
        $driveListBox.Location = New-Object System.Drawing.Point(20, 90)
        $driveListBox.Size = New-Object System.Drawing.Size(560, 150)
        $driveListBox.BackColor = [System.Drawing.Color]::Black
        $driveListBox.ForeColor = [System.Drawing.Color]::Lime
        $driveListBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $changeDriveForm.Controls.Add($driveListBox)

        # Populate drive list
        Add-Status "Getting list of drives..."
        try {
            $drives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name='Name';Expression={$_.DeviceID}},
                @{Name='VolumeName';Expression={$_.VolumeName}},
                @{Name='Size (GB)';Expression={[math]::round($_.Size/1GB, 0)}},
                @{Name='FreeSpace (GB)';Expression={[math]::round($_.FreeSpace/1GB, 0)}}

            foreach ($drive in $drives) {
                $driveInfo = "$($drive.Name) - $($drive.VolumeName) - Size: $($drive.'Size (GB)') GB - Free: $($drive.'FreeSpace (GB)') GB"
                $driveListBox.Items.Add($driveInfo)
            }

            if ($driveListBox.Items.Count -gt 0) {
                $driveListBox.SelectedIndex = 0
            }

            Add-Status "Found $($drives.Count) drives."
        } catch {
            Add-Status "Error getting drive list: $_"
        }

        # Old drive letter label
        $oldLetterLabel = New-Object System.Windows.Forms.Label
        $oldLetterLabel.Text = "Select Drive Letter to Change:"
        $oldLetterLabel.Location = New-Object System.Drawing.Point(20, 250)
        $oldLetterLabel.Size = New-Object System.Drawing.Size(200, 20)
        $oldLetterLabel.ForeColor = [System.Drawing.Color]::White
        $oldLetterLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $changeDriveForm.Controls.Add($oldLetterLabel)

        # Old drive letter textbox
        $oldLetterTextBox = New-Object System.Windows.Forms.TextBox
        $oldLetterTextBox.Location = New-Object System.Drawing.Point(230, 250)
        $oldLetterTextBox.Size = New-Object System.Drawing.Size(50, 20)
        $oldLetterTextBox.BackColor = [System.Drawing.Color]::Black
        $oldLetterTextBox.ForeColor = [System.Drawing.Color]::Lime
        $oldLetterTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $oldLetterTextBox.MaxLength = 1
        $changeDriveForm.Controls.Add($oldLetterTextBox)

        # New drive letter label
        $newLetterLabel = New-Object System.Windows.Forms.Label
        $newLetterLabel.Text = "New Drive Letter:"
        $newLetterLabel.Location = New-Object System.Drawing.Point(20, 290)
        $newLetterLabel.Size = New-Object System.Drawing.Size(200, 20)
        $newLetterLabel.ForeColor = [System.Drawing.Color]::White
        $newLetterLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $changeDriveForm.Controls.Add($newLetterLabel)

        # New drive letter textbox
        $newLetterTextBox = New-Object System.Windows.Forms.TextBox
        $newLetterTextBox.Location = New-Object System.Drawing.Point(230, 290)
        $newLetterTextBox.Size = New-Object System.Drawing.Size(50, 20)
        $newLetterTextBox.BackColor = [System.Drawing.Color]::Black
        $newLetterTextBox.ForeColor = [System.Drawing.Color]::Lime
        $newLetterTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $newLetterTextBox.MaxLength = 1
        $changeDriveForm.Controls.Add($newLetterTextBox)

        # Status textbox
        $changeStatusTextBox = New-Object System.Windows.Forms.TextBox
        $changeStatusTextBox.Multiline = $true
        $changeStatusTextBox.ScrollBars = "Vertical"
        $changeStatusTextBox.Location = New-Object System.Drawing.Point(20, 380)
        $changeStatusTextBox.Size = New-Object System.Drawing.Size(560, 70)
        $changeStatusTextBox.BackColor = [System.Drawing.Color]::Black
        $changeStatusTextBox.ForeColor = [System.Drawing.Color]::Lime
        $changeStatusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
        $changeStatusTextBox.ReadOnly = $true
        $changeStatusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $changeStatusTextBox.Text = "Ready to change drive letter..."
        $changeDriveForm.Controls.Add($changeStatusTextBox)

        # Function to add status message to the change form
        function Add-ChangeStatus {
            param([string]$message)
            $changeStatusTextBox.AppendText("$message`r`n")
            $changeStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }

        # Update old letter when drive is selected
        $driveListBox.Add_SelectedIndexChanged({
            if ($driveListBox.SelectedItem) {
                $selectedDrive = $driveListBox.SelectedItem.ToString()
                $driveLetter = $selectedDrive.Substring(0, 1)
                $oldLetterTextBox.Text = $driveLetter
            }
        })

        # Change button
        $changeButton = New-DynamicButton -text "Change Drive Letter" -x 20 -y 330 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            $oldLetter = $oldLetterTextBox.Text.Trim().ToUpper()
            $newLetter = $newLetterTextBox.Text.Trim().ToUpper()

            # Validate input
            if ($oldLetter -eq "") {
                Add-ChangeStatus "Error: Please select a drive letter to change."
                return
            }

            if ($newLetter -eq "") {
                Add-ChangeStatus "Error: Please enter a new drive letter."
                return
            }

            if ($oldLetter -eq $newLetter) {
                Add-ChangeStatus "Error: New drive letter must be different from the current one."
                return
            }

            # Check if new letter is already in use
            $existingDrives = Get-WmiObject Win32_LogicalDisk | Select-Object -ExpandProperty DeviceID
            if ($existingDrives -contains "$($newLetter):") {
                Add-ChangeStatus "Error: Drive letter $newLetter is already in use."
                return
            }

            # Create diskpart script
            $tempFile = [System.IO.Path]::GetTempFileName()
            $diskpartScript = @"
select volume $oldLetter
assign letter=$newLetter
"@
            Set-Content -Path $tempFile -Value $diskpartScript

            Add-ChangeStatus "Changing drive $oldLetter to $newLetter..."

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
                    Add-ChangeStatus "Successfully changed drive letter from $oldLetter to $newLetter."
                    Add-Status "Changed drive letter from $oldLetter to $newLetter."

                    # Refresh drive list
                    $driveListBox.Items.Clear()
                    $drives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name='Name';Expression={$_.DeviceID}},
                        @{Name='VolumeName';Expression={$_.VolumeName}},
                        @{Name='Size (GB)';Expression={[math]::round($_.Size/1GB, 0)}},
                        @{Name='FreeSpace (GB)';Expression={[math]::round($_.FreeSpace/1GB, 0)}}

                    foreach ($drive in $drives) {
                        $driveInfo = "$($drive.Name) - $($drive.VolumeName) - Size: $($drive.'Size (GB)') GB - Free: $($drive.'FreeSpace (GB)') GB"
                        $driveListBox.Items.Add($driveInfo)
                    }

                    if ($driveListBox.Items.Count -gt 0) {
                        $driveListBox.SelectedIndex = 0
                    }
                } else {
                    Add-ChangeStatus "Error changing drive letter. Exit code: $($process.ExitCode)"
                }
            } catch {
                Add-ChangeStatus "Error: $_"
            } finally {
                # Clean up temp file
                if (Test-Path $tempFile) {
                    Remove-Item $tempFile -Force
                }
            }
        }
        $changeDriveForm.Controls.Add($changeButton)

        # Cancel button
        $cancelButton = New-DynamicButton -text "Cancel" -x 240 -y 330 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
            $changeDriveForm.Close()
        }
        $changeDriveForm.Controls.Add($cancelButton)

        # Show the form
        Add-Status "Opening Change Drive Letter dialog..."
        $changeDriveForm.ShowDialog()
    }
    $volumeForm.Controls.Add($btnChangeDriveLetter)

    # Shrink Volume button
    $btnShrinkVolume = New-DynamicButton -text "[2] Shrink Volume" -x 50 -y 130 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        Add-Status "Opening Disk Management to shrink volume..."
        try {
            # Open Disk Management
            Start-Process "diskmgmt.msc" -Verb RunAs
            Add-Status "Disk Management has been opened. Right-click on a volume and select 'Shrink Volume'."
        } catch {
            Add-Status "Error opening Disk Management: $_"
        }
    }
    $volumeForm.Controls.Add($btnShrinkVolume)

    # Extend Volume button
    $btnExtendVolume = New-DynamicButton -text "[3] Extend Volume" -x 50 -y 180 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        Add-Status "Opening Disk Management to extend volume..."
        try {
            # Open Disk Management
            Start-Process "diskmgmt.msc" -Verb RunAs
            Add-Status "Disk Management has been opened. Right-click on a volume and select 'Extend Volume'."
        } catch {
            Add-Status "Error opening Disk Management: $_"
        }
    }
    $volumeForm.Controls.Add($btnExtendVolume)

    # Rename Volume button
    $btnRenameVolume = New-DynamicButton -text "[4] Rename" -x 50 -y 230 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        Add-Status "Opening File Explorer to rename volumes..."
        try {
            # Open File Explorer to This PC
            Start-Process "explorer.exe" -ArgumentList "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
            Add-Status "File Explorer has been opened. Right-click on a drive and select 'Rename'."
        } catch {
            Add-Status "Error opening File Explorer: $_"
        }
    }
    $volumeForm.Controls.Add($btnRenameVolume)

    # Return to Main Menu button
    $btnReturn = New-DynamicButton -text "[0] Return to Menu" -x 50 -y 280 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        $volumeForm.Close()
    }
    $volumeForm.Controls.Add($btnReturn)

    # Show the form
    $volumeForm.ShowDialog()
}

# Activate Windows 10 Pro and Office 2019 Pro Plus
$buttonActivate = New-DynamicButton -text "Activate" -x 30 -y 420 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Create activation form
    $activateForm = New-Object System.Windows.Forms.Form
    $activateForm.Text = "Activation Options"
    $activateForm.Size = New-Object System.Drawing.Size(500, 450)
    $activateForm.StartPosition = "CenterScreen"
    $activateForm.BackColor = [System.Drawing.Color]::Black
    $activateForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $activateForm.MaximizeBox = $false
    $activateForm.MinimizeBox = $false

    # Add a gradient background to activation form
    $activateForm.Paint = {
        $graphics = $_.Graphics
        $rect = New-Object System.Drawing.Rectangle(0, 0, $activateForm.Width, $activateForm.Height)
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $rect,
            [System.Drawing.Color]::FromArgb(0, 0, 0),  # Black at top
            [System.Drawing.Color]::FromArgb(0, 30, 0), # Dark green at bottom
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
        )
        $graphics.FillRectangle($brush, $rect)
        $brush.Dispose()
    }

    # Title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Activating Windows 10 Pro / Office 2019 Pro Plus"
    $titleLabel.Location = New-Object System.Drawing.Point(-10, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(500, 30)
    $titleLabel.ForeColor = [System.Drawing.Color]::Lime
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $activateForm.Controls.Add($titleLabel)

    # Add animation to the title
    $titleTimer = New-Object System.Windows.Forms.Timer
    $titleTimer.Interval = 800
    $titleTimer.Add_Tick({
        if ($titleLabel.ForeColor -eq [System.Drawing.Color]::Lime) {
            $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 200, 0)
        } else {
            $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        }
    })
    $titleTimer.Start()

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(12, 280)
    $statusTextBox.Size = New-Object System.Drawing.Size(460, 120)
    $statusTextBox.BackColor = [System.Drawing.Color]::Black
    $statusTextBox.ForeColor = [System.Drawing.Color]::Lime
    $statusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $statusTextBox.ReadOnly = $true
    $statusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $statusTextBox.Text = "Status messages will appear here..."
    $activateForm.Controls.Add($statusTextBox)

    # Function to add status message
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

    # Activation buttons
    $btnWin10Pro = New-DynamicButton -text "Active Windows 10 Pro" -x 12 -y 70 -width 460 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        try {
            Add-Status "Checking Activation Status of Windows..."
            $windowsStatus = & cscript //nologo "$env:windir\system32\slmgr.vbs" /dli
            $isWindowsActivated = $windowsStatus -match "License Status: Licensed"

            if ($isWindowsActivated) {
                Add-Status "Windows activated."
                return
            }

            Add-Status "Windows not activated. Activating Windows 10 Pro..."
            $command = "slmgr /ipk R84N4-RPC7Q-W8TKM-VM7Y4-7H66Y && slmgr /ato"

            # Create a process to run the command with elevated privileges
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-Command Start-Process cmd.exe -ArgumentList '/c $command' -Verb RunAs -WindowStyle Hidden"
            $psi.UseShellExecute = $true
            $psi.Verb = "runas"
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

            # Start the process
            [System.Diagnostics.Process]::Start($psi)

            Add-Status "Starting activation process for Windows 10 Pro."
        } catch {
            Add-Status "Lỗi khi kích hoạt Windows: $_"
        }
    }
    $activateForm.Controls.Add($btnWin10Pro)

    # Add button to activate Office 2019
    $btnOffice = New-DynamicButton -text "Active Office2019ProPlus" -x 12 -y 120 -width 460 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        try {
            Add-Status "Checking Activation Status of Office..."

            # Check if Office16 path exists
            $office16Path = "C:\Program Files\Microsoft Office\Office16\ospp.vbs"
            $office15Path = "C:\Program Files\Microsoft Office\Office15\ospp.vbs"

            if (Test-Path $office16Path) {
                $officePath = $office16Path
                Add-Status "Found Office16."
            } elseif (Test-Path $office15Path) {
                $officePath = $office15Path
                Add-Status "Found Office15."
            } else {
                Add-Status "Not found Office16 or Office15. Using Office16 path by default."
                $officePath = $office16Path
            }

            # Kiểm tra trạng thái kích hoạt Office
            $officeStatus = & cscript //nologo "$officePath" /dstatus
            $isOfficeActivated = $officeStatus -match "LICENSE STATUS:  ---LICENSED---"

            if ($isOfficeActivated) {
                Add-Status "Office activated."
                return
            }

            Add-Status "Office not activated. Activating Office 2019 Pro Plus..."
            $command = "cscript `"$officePath`" /inpkey:Q2NKY-J42YJ-X2KVK-9Q9PT-MKP63 && cscript `"$officePath`" /act"

            # Create a process to run the command with elevated privileges
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-Command Start-Process cmd.exe -ArgumentList '/c $command' -Verb RunAs -WindowStyle Hidden"
            $psi.UseShellExecute = $true
            $psi.Verb = "runas"
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

            # Start the process
            [System.Diagnostics.Process]::Start($psi)

            Add-Status "Starting activation process for Office 2019."
        } catch {
            Add-Status "Error activating Office: $_"
        }
    }
    $activateForm.Controls.Add($btnOffice)

    # Add button to upgrade Windows 10 Home to Pro
    $btnWin10Home = New-DynamicButton -text "Win10Home to Win10Pro" -x 12 -y 170 -width 460 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        try {
            Add-Status "Checking Windows version..."

            # Kiểm tra phiên bản Windows
            $windowsEdition = (Get-WmiObject -Class Win32_OperatingSystem).Caption

            if ($windowsEdition -match "Pro") {
                Add-Status "Device is already running Windows 10 Pro."
                return
            }

            if (-not ($windowsEdition -match "Home")) {
                Add-Status "Device is not running Windows 10 Home. Cannot upgrade to Pro using this method."
                return
            }

            Add-Status "Upgrading Windows 10 Home to Pro..."
            $command = "sc config LicenseManager start= auto & net start LicenseManager & sc config wuauserv start= auto & net start wuauserv & changepk.exe /productkey VK7JG-NPHTM-C97JM-9MPGT-3V66T"

            # Create a process to run the command with elevated privileges
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-Command Start-Process cmd.exe -ArgumentList '/c $command' -Verb RunAs -WindowStyle Hidden"
            $psi.UseShellExecute = $true
            $psi.Verb = "runas"
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

            # Start the process
            [System.Diagnostics.Process]::Start($psi)

            Add-Status "Starting upgrade process for Windows 10 Home to Pro."
        } catch {
            Add-Status "Error upgrading Windows: $_"
        }
    }
    $activateForm.Controls.Add($btnWin10Home)

    # Return to Main Menu button
    $btnReturn = New-DynamicButton -text "[0] Return to Menu" -x 12 -y 220 -width 460 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        $activateForm.Close()
    }
    $activateForm.Controls.Add($btnReturn)

    # Show the form
    $activateForm.ShowDialog()
}

# Edit Features
$buttonTurnOnFeatures = New-DynamicButton -text "Turn On Features" -x 430 -y 100 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Create Windows Features form
    $featuresForm = New-Object System.Windows.Forms.Form
    $featuresForm.Text = "Windows Features"
    $featuresForm.Size = New-Object System.Drawing.Size(500, 500)
    $featuresForm.StartPosition = "CenterScreen"
    $featuresForm.BackColor = [System.Drawing.Color]::Black
    $featuresForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $featuresForm.MaximizeBox = $false
    $featuresForm.MinimizeBox = $false

    # Add a gradient background
    $featuresForm.Paint = {
        $graphics = $_.Graphics
        $rect = New-Object System.Drawing.Rectangle(0, 0, $featuresForm.Width, $featuresForm.Height)
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $rect,
            [System.Drawing.Color]::FromArgb(0, 0, 0),  # Black at top
            [System.Drawing.Color]::FromArgb(0, 40, 0), # Dark green at bottom
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
        )
        $graphics.FillRectangle($brush, $rect)
        $brush.Dispose()
    }

    # Title label with animation
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "WINDOWS FEATURES MANAGEMENT"
    $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(500, 40)
    $titleLabel.ForeColor = [System.Drawing.Color]::Lime
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $titleLabel.Padding = New-Object System.Windows.Forms.Padding(5)

    # Add animation to the title
    $titleTimer = New-Object System.Windows.Forms.Timer
    $titleTimer.Interval = 500
    $titleTimer.Add_Tick({
        if ($titleLabel.ForeColor -eq [System.Drawing.Color]::Lime) {
            $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 220, 0)
        } else {
            $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        }
    })
    $titleTimer.Start()

    $featuresForm.Controls.Add($titleLabel)

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(50, 330)
    $statusTextBox.Size = New-Object System.Drawing.Size(400, 120)
    $statusTextBox.BackColor = [System.Drawing.Color]::Black
    $statusTextBox.ForeColor = [System.Drawing.Color]::Lime
    $statusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $statusTextBox.ReadOnly = $true
    $featuresForm.Controls.Add($statusTextBox)

    # Function to add status message
    function Add-Status {
        param([string]$message)
        $statusTextBox.AppendText("$message`r`n")
        $statusTextBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Status text box with improved styling
    $statusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $statusTextBox.Text = "Status messages will appear here..."

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

    # Enable .NET Framework 3.5 button
    $btnEnableNetFx = New-DynamicButton -text "Enable .NET Framework 3.5" -x 50 -y 80 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        Add-Status "Checking .NET Framework 3.5 status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:NetFx3"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Disabled") {
                Add-Status "Enabling .NET Framework 3.5..."
                $enableCmd = "dism /online /enable-feature /featurename:NetFx3 /all /norestart"
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $enableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -WindowStyle Hidden -Wait
                Add-Status ".NET Framework 3.5 has been enabled."
            } else {
                Add-Status ".NET Framework 3.5 is already enabled."
            }
        } catch {
            Add-Status "Error: $_"
        }
    }
    $featuresForm.Controls.Add($btnEnableNetFx)

    # Enable WCF-HTTP-Activation button
    $btnEnableWcfHttp = New-DynamicButton -text "Enable WCF-HTTP-Activation" -x 50 -y 130 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        Add-Status "Checking WCF-HTTP-Activation status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:WCF-HTTP-Activation"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Disabled") {
                Add-Status "Enabling WCF-HTTP-Activation..."
                $enableCmd = "DISM /Online /Enable-Feature /FeatureName:WCF-HTTP-Activation /All /Quiet /NoRestart"
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $enableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -WindowStyle Hidden -Wait
                Add-Status "WCF-HTTP-Activation has been enabled."
            } else {
                Add-Status "WCF-HTTP-Activation is already enabled."
            }
        } catch {
            Add-Status "Error: $_"
        }
    }
    $featuresForm.Controls.Add($btnEnableWcfHttp)

    # Enable WCF-NonHTTP-Activation button
    $btnEnableWcfNonHttp = New-DynamicButton -text "Enable WCF-NonHTTP-Activation" -x 50 -y 180 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        Add-Status "Checking WCF-NonHTTP-Activation status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:WCF-NonHTTP-Activation"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Disabled") {
                Add-Status "Enabling WCF-NonHTTP-Activation..."
                $enableCmd = "DISM /Online /Enable-Feature /FeatureName:WCF-NonHTTP-Activation /All /Quiet /NoRestart"
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $enableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -WindowStyle Hidden -Wait
                Add-Status "WCF-NonHTTP-Activation has been enabled."
            } else {
                Add-Status "WCF-NonHTTP-Activation is already enabled."
            }
        } catch {
            Add-Status "Error: $_"
        }
    }
    $featuresForm.Controls.Add($btnEnableWcfNonHttp)

    # Disable Internet Explorer 11 button
    $btnDisableIE = New-DynamicButton -text "Disable Internet Explorer 11" -x 50 -y 230 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        Add-Status "Checking Internet Explorer 11 status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:Internet-Explorer-Optional-amd64"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Enabled") {
                Add-Status "Disabling Internet Explorer 11..."
                $disableCmd = "dism /online /disable-feature /featurename:Internet-Explorer-Optional-amd64 /norestart"
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $disableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -WindowStyle Hidden -Wait
                Add-Status "Internet Explorer 11 has been disabled."
            } else {
                Add-Status "Internet Explorer 11 is already disabled."
            }
        } catch {
            Add-Status "Error: $_"
        }
    }
    $featuresForm.Controls.Add($btnDisableIE)

    # Return to Main Menu button
    $btnReturn = New-DynamicButton -text "Return to Main Menu" -x 50 -y 280 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        $featuresForm.Close()
    }
    $featuresForm.Controls.Add($btnReturn)

    # Show the form
    $featuresForm.ShowDialog()
}

# Rename Device
$buttonRenameDevice = New-DynamicButton -text "Rename Device" -x 430 -y 180 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Create device rename form
    $renameForm = New-Object System.Windows.Forms.Form
    $renameForm.Text = "Rename Device"
    $renameForm.Size = New-Object System.Drawing.Size(500, 300)
    $renameForm.StartPosition = "CenterScreen"
    $renameForm.BackColor = [System.Drawing.Color]::Black
    $renameForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $renameForm.MaximizeBox = $false
    $renameForm.MinimizeBox = $false

    # Create title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Rename Current Device"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.Size = New-Object System.Drawing.Size(480, 40)
    $titleLabel.Location = New-Object System.Drawing.Point(10, 20)
    $renameForm.Controls.Add($titleLabel)

    # Get current computer name
    $currentName = $env:COMPUTERNAME

    # Current device name label
    $currentLabel = New-Object System.Windows.Forms.Label
    $currentLabel.Text = "Current Device Name: $currentName"
    $currentLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $currentLabel.ForeColor = [System.Drawing.Color]::White
    $currentLabel.Size = New-Object System.Drawing.Size(480, 30)
    $currentLabel.Location = New-Object System.Drawing.Point(20, 70)
    $renameForm.Controls.Add($currentLabel)

    # New name label
    $newNameLabel = New-Object System.Windows.Forms.Label
    $newNameLabel.Text = "New Device Name:"
    $newNameLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $newNameLabel.ForeColor = [System.Drawing.Color]::White
    $newNameLabel.Size = New-Object System.Drawing.Size(150, 30)
    $newNameLabel.Location = New-Object System.Drawing.Point(20, 120)
    $renameForm.Controls.Add($newNameLabel)

    # New name textbox
    $newNameTextBox = New-Object System.Windows.Forms.TextBox
    $newNameTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $newNameTextBox.Size = New-Object System.Drawing.Size(300, 30)
    $newNameTextBox.Location = New-Object System.Drawing.Point(170, 120)
    $newNameTextBox.BackColor = [System.Drawing.Color]::Black
    $newNameTextBox.ForeColor = [System.Drawing.Color]::Green
    $newNameTextBox.Text = $currentName
    $renameForm.Controls.Add($newNameTextBox)

    # Rename button
    $renameButton = New-Object System.Windows.Forms.Button
    $renameButton.Text = "Rename Device"
    $renameButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $renameButton.ForeColor = [System.Drawing.Color]::White
    $renameButton.BackColor = [System.Drawing.Color]::FromArgb(0, 180, 0)
    $renameButton.Size = New-Object System.Drawing.Size(200, 40)
    $renameButton.Location = New-Object System.Drawing.Point(30, 180)
    $renameButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $renameButton.Add_Click({
        $newName = $newNameTextBox.Text.Trim()

        # Validate new name
        if ($newName -eq "") {
            [System.Windows.Forms.MessageBox]::Show("Device name cannot be empty.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Check if name is valid (15 characters max, no special characters)
        if ($newName.Length -gt 15) {
            [System.Windows.Forms.MessageBox]::Show("Device name must be 15 characters or less.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        if ($newName -match '[^\w.-]') {
            [System.Windows.Forms.MessageBox]::Show("Device name contains invalid characters. Use only letters, numbers, hyphens, and periods.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Check if name is the same as current
        if ($newName -eq $currentName) {
            [System.Windows.Forms.MessageBox]::Show("New name is the same as current name. No changes needed.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        try {
            # Create a command to rename the computer
            $command = "Rename-Computer -NewName '$newName' -Force -Restart"

            # Create a process to run the command with elevated privileges
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-Command Start-Process powershell.exe -ArgumentList '-Command $command' -Verb RunAs"
            $psi.UseShellExecute = $true
            $psi.Verb = "runas"

            # Start the process
            [System.Diagnostics.Process]::Start($psi)

            # Show success message
            [System.Windows.Forms.MessageBox]::Show("Device rename command has been initiated. If prompted, please allow the elevation request. Your computer will restart to apply the changes.", "Device Rename", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $renameForm.Close()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error renaming device: $_`n`nNote: This operation requires administrative privileges.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $renameForm.Controls.Add($renameButton)

    # Cancel button
    $cancelButton = New-RedButton -text "Cancel" -x 250 -y 180 -width 200 -height 40 -clickAction {
        $renameForm.Close()
    }
    $renameForm.Controls.Add($cancelButton)

    # Show the form
    $renameForm.ShowDialog()
}

# Set Password
$buttonSetPassword = New-DynamicButton -text "Set Password" -x 430 -y 260 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Create password setting form
    $passwordForm = New-Object System.Windows.Forms.Form
    $passwordForm.Text = "Set Password"
    $passwordForm.Size = New-Object System.Drawing.Size(500, 350)
    $passwordForm.StartPosition = "CenterScreen"
    $passwordForm.BackColor = [System.Drawing.Color]::Black
    $passwordForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $passwordForm.MaximizeBox = $false
    $passwordForm.MinimizeBox = $false

    # Create title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Set Password for Current User"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.Size = New-Object System.Drawing.Size(480, 40)
    $titleLabel.Location = New-Object System.Drawing.Point(10, 20)
    $passwordForm.Controls.Add($titleLabel)

    # Current user label
    $currentUser = $env:USERNAME
    $userLabel = New-Object System.Windows.Forms.Label
    $userLabel.Text = "Current User: $currentUser"
    $userLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $userLabel.ForeColor = [System.Drawing.Color]::White
    $userLabel.Size = New-Object System.Drawing.Size(480, 30)
    $userLabel.Location = New-Object System.Drawing.Point(20, 70)
    $passwordForm.Controls.Add($userLabel)

    # Password label
    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Text = "New Password:"
    $passwordLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $passwordLabel.ForeColor = [System.Drawing.Color]::White
    $passwordLabel.Size = New-Object System.Drawing.Size(150, 30)
    $passwordLabel.Location = New-Object System.Drawing.Point(20, 120)
    $passwordForm.Controls.Add($passwordLabel)

    # Password textbox
    $passwordTextBox = New-Object System.Windows.Forms.TextBox
    $passwordTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $passwordTextBox.Size = New-Object System.Drawing.Size(300, 30)
    $passwordTextBox.Location = New-Object System.Drawing.Point(170, 120)
    $passwordTextBox.UseSystemPasswordChar = $true
    $passwordTextBox.BackColor = [System.Drawing.Color]::Black
    $passwordTextBox.ForeColor = [System.Drawing.Color]::Green
    $passwordForm.Controls.Add($passwordTextBox)

    # Confirm password label
    $confirmLabel = New-Object System.Windows.Forms.Label
    $confirmLabel.Text = "Confirm Password:"
    $confirmLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $confirmLabel.ForeColor = [System.Drawing.Color]::White
    $confirmLabel.Size = New-Object System.Drawing.Size(150, 30)
    $confirmLabel.Location = New-Object System.Drawing.Point(20, 170)
    $passwordForm.Controls.Add($confirmLabel)

    # Confirm password textbox
    $confirmTextBox = New-Object System.Windows.Forms.TextBox
    $confirmTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $confirmTextBox.Size = New-Object System.Drawing.Size(300, 30)
    $confirmTextBox.Location = New-Object System.Drawing.Point(170, 170)
    $confirmTextBox.UseSystemPasswordChar = $true
    $confirmTextBox.BackColor = [System.Drawing.Color]::Black
    $confirmTextBox.ForeColor = [System.Drawing.Color]::Green
    $passwordForm.Controls.Add($confirmTextBox)

    # Set Password button
    $setButton = New-Object System.Windows.Forms.Button
    $setButton.Text = "Set Password"
    $setButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $setButton.ForeColor = [System.Drawing.Color]::White
    $setButton.BackColor = [System.Drawing.Color]::FromArgb(0, 180, 0)
    $setButton.Size = New-Object System.Drawing.Size(200, 40)
    $setButton.Location = New-Object System.Drawing.Point(30, 230)
    $setButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $setButton.Add_Click({
        $password = $passwordTextBox.Text
        $confirm = $confirmTextBox.Text

        # Validate password
        if ($password -eq "") {
            [System.Windows.Forms.MessageBox]::Show("Password cannot be empty.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Check if passwords match
        if ($password -ne $confirm) {
            [System.Windows.Forms.MessageBox]::Show("Passwords do not match.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        try {
            # Create a command to set the password
            $command = "net user $currentUser $password"

            # Create a process to run the command with elevated privileges
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-Command Start-Process cmd.exe -ArgumentList '/c $command' -Verb RunAs -WindowStyle Hidden"
            $psi.UseShellExecute = $true
            $psi.Verb = "runas"
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

            # Start the process
            [System.Diagnostics.Process]::Start($psi)

            # Show success message
            [System.Windows.Forms.MessageBox]::Show("Password change command has been initiated. If prompted, please allow the elevation request.", "Password Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $passwordForm.Close()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error setting password: $_`n`nNote: This operation requires administrative privileges.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $passwordForm.Controls.Add($setButton)

    # Cancel button
    $cancelButton = New-RedButton -text "Cancel" -x 250 -y 230 -width 200 -height 40 -clickAction {
        $passwordForm.Close()
    }
    $passwordForm.Controls.Add($cancelButton)

    # Show the form
    $passwordForm.ShowDialog()
}

# Join Domain
$buttonJoinDomain = New-DynamicButton -text "Join Domain" -x 430 -y 340 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Create domain/workgroup join form
    $joinForm = New-Object System.Windows.Forms.Form
    $joinForm.Text = "Join Domain/Workgroup"
    $joinForm.Size = New-Object System.Drawing.Size(500, 450)
    $joinForm.StartPosition = "CenterScreen"
    $joinForm.BackColor = [System.Drawing.Color]::Black
    $joinForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $joinForm.MaximizeBox = $false
    $joinForm.MinimizeBox = $false

    # Create title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Join Domain or Workgroup"
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

    $radioDomain = New-Object System.Windows.Forms.RadioButton
    $radioDomain.Text = "Join Domain"
    $radioDomain.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioDomain.ForeColor = [System.Drawing.Color]::White
    $radioDomain.Location = New-Object System.Drawing.Point(20, 30)
    $radioDomain.Size = New-Object System.Drawing.Size(200, 30)
    $radioDomain.BackColor = [System.Drawing.Color]::Black
    $radioDomain.Checked = $true

    $radioWorkgroup = New-Object System.Windows.Forms.RadioButton
    $radioWorkgroup.Text = "Join Workgroup"
    $radioWorkgroup.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioWorkgroup.ForeColor = [System.Drawing.Color]::White
    $radioWorkgroup.Location = New-Object System.Drawing.Point(240, 30)
    $radioWorkgroup.Size = New-Object System.Drawing.Size(200, 30)
    $radioWorkgroup.BackColor = [System.Drawing.Color]::Black
    $radioWorkgroup.Checked = $false

    $groupBox.Controls.Add($radioDomain)
    $groupBox.Controls.Add($radioWorkgroup)
    $joinForm.Controls.Add($groupBox)

    # Name label
    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = "Domain Name:"
    $nameLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $nameLabel.ForeColor = [System.Drawing.Color]::White
    $nameLabel.Size = New-Object System.Drawing.Size(150, 30)
    $nameLabel.Location = New-Object System.Drawing.Point(20, 230)
    $joinForm.Controls.Add($nameLabel)

    # Name textbox
    $nameTextBox = New-Object System.Windows.Forms.TextBox
    $nameTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $nameTextBox.Size = New-Object System.Drawing.Size(300, 30)
    $nameTextBox.Location = New-Object System.Drawing.Point(170, 230)
    $nameTextBox.BackColor = [System.Drawing.Color]::Black
    $nameTextBox.ForeColor = [System.Drawing.Color]::Green
    $nameTextBox.Text = ""
    $joinForm.Controls.Add($nameTextBox)

    # Username label (for domain)
    $usernameLabel = New-Object System.Windows.Forms.Label
    $usernameLabel.Text = "Username:"
    $usernameLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $usernameLabel.ForeColor = [System.Drawing.Color]::White
    $usernameLabel.Size = New-Object System.Drawing.Size(150, 30)
    $usernameLabel.Location = New-Object System.Drawing.Point(20, 270)
    $usernameLabel.Visible = $true
    $joinForm.Controls.Add($usernameLabel)

    # Username textbox
    $usernameTextBox = New-Object System.Windows.Forms.TextBox
    $usernameTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $usernameTextBox.Size = New-Object System.Drawing.Size(300, 30)
    $usernameTextBox.Location = New-Object System.Drawing.Point(170, 270)
    $usernameTextBox.BackColor = [System.Drawing.Color]::Black
    $usernameTextBox.ForeColor = [System.Drawing.Color]::Green
    $usernameTextBox.Visible = $true
    $joinForm.Controls.Add($usernameTextBox)

    # Password label (for domain)
    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Text = "Password:"
    $passwordLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $passwordLabel.ForeColor = [System.Drawing.Color]::White
    $passwordLabel.Size = New-Object System.Drawing.Size(150, 30)
    $passwordLabel.Location = New-Object System.Drawing.Point(20, 310)
    $passwordLabel.Visible = $true
    $joinForm.Controls.Add($passwordLabel)

    # Password textbox
    $passwordTextBox = New-Object System.Windows.Forms.TextBox
    $passwordTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $passwordTextBox.Size = New-Object System.Drawing.Size(300, 30)
    $passwordTextBox.Location = New-Object System.Drawing.Point(170, 310)
    $passwordTextBox.BackColor = [System.Drawing.Color]::Black
    $passwordTextBox.ForeColor = [System.Drawing.Color]::Green
    $passwordTextBox.UseSystemPasswordChar = $true
    $passwordTextBox.Visible = $true
    $joinForm.Controls.Add($passwordTextBox)

    # Event handlers for radio buttons
    $radioDomain.Add_CheckedChanged({
        if ($radioDomain.Checked) {
            $nameLabel.Text = "Domain Name:"
            $nameTextBox.Text = ""
            $usernameLabel.Visible = $true
            $usernameTextBox.Visible = $true
            $passwordLabel.Visible = $true
            $passwordTextBox.Visible = $true
            $joinButton.Location = New-Object System.Drawing.Point(30, 350)
            $cancelButton.Location = New-Object System.Drawing.Point(250, 350)
            $joinForm.Size = New-Object System.Drawing.Size(500, 450)
        }
    })

    $radioWorkgroup.Add_CheckedChanged({
        if ($radioWorkgroup.Checked) {
            $nameLabel.Text = "Workgroup Name:"
            $nameTextBox.Text = "WORKGROUP"
            $usernameLabel.Visible = $false
            $usernameTextBox.Visible = $false
            $passwordLabel.Visible = $false
            $passwordTextBox.Visible = $false
            $joinButton.Location = New-Object System.Drawing.Point(30, 280)
            $cancelButton.Location = New-Object System.Drawing.Point(250, 280)
            $joinForm.Size = New-Object System.Drawing.Size(500, 380)
        }
    })

    # Join button
    $joinButton = New-DynamicButton -text "Join" -x 30 -y 350 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 180, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 220, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 140, 0)) -textColor ([System.Drawing.Color]::White) -fontSize 12 -fontStyle ([System.Drawing.FontStyle]::Bold)
    $joinButton.Add_Click({
        $name = $nameTextBox.Text.Trim()

        # Validate name
        if ($name -eq "") {
            [System.Windows.Forms.MessageBox]::Show("Name cannot be empty.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        try {
            if ($radioDomain.Checked) {
                # Join domain
                $username = $usernameTextBox.Text.Trim()
                $password = $passwordTextBox.Text

                if ($username -eq "" -or $password -eq "") {
                    [System.Windows.Forms.MessageBox]::Show("Username and password are required for domain join.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }

                # Create a command to join domain
                $command = "Add-Computer -DomainName '$name' -Credential (New-Object System.Management.Automation.PSCredential ('$username', (ConvertTo-SecureString '$password' -AsPlainText -Force))) -Restart -Force"

                # Create a process to run the command with elevated privileges
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = "powershell.exe"
                $psi.Arguments = "-Command Start-Process powershell.exe -ArgumentList '-Command $command' -Verb RunAs"
                $psi.UseShellExecute = $true
                $psi.Verb = "runas"

                # Start the process
                [System.Diagnostics.Process]::Start($psi)

                # Show success message
                [System.Windows.Forms.MessageBox]::Show("Domain join command has been initiated. If prompted, please allow the elevation request. Your computer will restart to apply the changes.", "Domain Join", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $joinForm.Close()
            }
            else {
                # Join workgroup
                $command = "Add-Computer -WorkgroupName '$name' -Restart -Force"

                # Create a process to run the command with elevated privileges
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = "powershell.exe"
                $psi.Arguments = "-Command Start-Process powershell.exe -ArgumentList '-Command $command' -Verb RunAs"
                $psi.UseShellExecute = $true
                $psi.Verb = "runas"

                # Start the process
                [System.Diagnostics.Process]::Start($psi)

                # Show success message
                [System.Windows.Forms.MessageBox]::Show("Workgroup join command has been initiated. If prompted, please allow the elevation request. Your computer will restart to apply the changes.", "Workgroup Join", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $joinForm.Close()
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error joining domain/workgroup: $_`n`nNote: This operation requires administrative privileges.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $joinForm.Controls.Add($joinButton)

    # Cancel button
    $cancelButton = New-DynamicButton -text "Cancel" -x 250 -y 350 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        $joinForm.Close()
    }
    $joinForm.Controls.Add($cancelButton)

    # Show the form
    $joinForm.ShowDialog()
}

# Exit button
$buttonExit = New-DynamicButton -text "Exit" -x 430 -y 420 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
    $form.Close()
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
