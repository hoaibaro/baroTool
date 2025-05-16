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

# Run All Options
$buttonRunAll = New-DynamicButton -text "[1] Run All Options" -x 30 -y 100 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Hide the main menu
    Hide-MainMenu
}

# Install All Software
$buttonInstallSoftware = New-DynamicButton -text "[2] Install All Software" -x 30 -y 180 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Hide the main menu
    Hide-MainMenu
    # Create device type selection form
    $deviceTypeForm = New-Object System.Windows.Forms.Form
    $deviceTypeForm.Text = "Select Device Type"
    $deviceTypeForm.Size = New-Object System.Drawing.Size(500, 400)
    $deviceTypeForm.StartPosition = "CenterScreen"
    $deviceTypeForm.BackColor = [System.Drawing.Color]::Black
    $deviceTypeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $deviceTypeForm.MaximizeBox = $false
    $deviceTypeForm.MinimizeBox = $false

    # Add a gradient background
    $deviceTypeForm.Paint = {
        $graphics = $_.Graphics
        $rect = New-Object System.Drawing.Rectangle(0, 0, $deviceTypeForm.Width, $deviceTypeForm.Height)
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
    $titleLabel.Text = "SELECT DEVICE TYPE"
    $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(500, 40)
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

    $deviceTypeForm.Controls.Add($titleLabel)

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(50, 250)
    $statusTextBox.Size = New-Object System.Drawing.Size(400, 100)
    $statusTextBox.BackColor = [System.Drawing.Color]::Black
    $statusTextBox.ForeColor = [System.Drawing.Color]::Lime
    $statusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $statusTextBox.ReadOnly = $true
    $statusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $statusTextBox.Text = "Please select a device type..."
    $deviceTypeForm.Controls.Add($statusTextBox)

    # Function to add status message
    function Add-Status {
        param([string]$message)

        # Clear placeholder text on first message
        if ($statusTextBox.Text -eq "Please select a device type...") {
            $statusTextBox.Clear()
        }

        # Add timestamp to message
        $timestamp = Get-Date -Format "HH:mm:ss"
        $statusTextBox.AppendText("[$timestamp] $message`r`n")
        $statusTextBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Desktop button
    $btnDesktop = New-DynamicButton -text "DESKTOP" -x 100 -y 80 -width 300 -height 70 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        Add-Status "Starting desktop software installation process..."

        # First copy all necessary files
        Add-Status "Step 1: Copying required files..."
        $copyResult = Copy-SoftwareFiles -deviceType "Desktop"

        if ($copyResult) {
            # Then install the software
            Add-Status "Step 2: Installing software..."
            $installResult = Install-Software -deviceType "Desktop"

            if ($installResult) {
                Add-Status "Desktop software installation completed successfully!"
            }
            else {
                Add-Status "Warning: Some software installations may have failed. Check the log for details."
            }
        }
        else {
            Add-Status "Error: Failed to copy required files. Installation aborted."
        }
    }
    $deviceTypeForm.Controls.Add($btnDesktop)

    # Laptop button
    $btnLaptop = New-DynamicButton -text "LAPTOP" -x 100 -y 170 -width 300 -height 70 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        Add-Status "Starting laptop software installation process..."

        # First copy all necessary files
        Add-Status "Step 1: Copying required files..."
        $copyResult = Copy-SoftwareFiles -deviceType "Laptop"

        if ($copyResult) {
            # Then install the software
            Add-Status "Step 2: Installing software..."
            $installResult = Install-Software -deviceType "Laptop"

            if ($installResult) {
                Add-Status "Laptop software installation completed successfully!"
            }
            else {
                Add-Status "Warning: Some software installations may have failed. Check the log for details."
            }
        }
        else {
            Add-Status "Error: Failed to copy required files. Installation aborted."
        }
    }
    $deviceTypeForm.Controls.Add($btnLaptop)

    # When the form is closed, show the main menu again
    $deviceTypeForm.Add_FormClosed({
        Show-MainMenu
    })

    # Show the form
    $deviceTypeForm.ShowDialog()

    # Function to copy files
    function Copy-SoftwareFiles {
        param (
            [string]$deviceType # "Desktop" or "Laptop"
        )

        try {
            # Create temp directory
            $tempDir = "$env:USERPROFILE\Downloads\SETUP"
            if (-not (Test-Path $tempDir)) {
                Add-Status "Creating temporary folder..."
                New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
                Add-Status "Temporary folder created successfully!"
            }
            else {
                Add-Status "Temporary folder already exists. Skipping..."
            }

            # Copy setup files
            if (-not (Test-Path "$tempDir\Software")) {
                Add-Status "Copying setup files..."
                $setupSource = "D:\SOFTWARE\PAYOO\SETUP"
                if (Test-Path $setupSource) {
                    Copy-Item -Path $setupSource -Destination "$tempDir\Software" -Recurse -Force
                    Add-Status "SetupFiles has been copied successfully!"
                }
                else {
                    Add-Status "Warning: Setup source folder not found at $setupSource"
                }
            }
            else {
                Add-Status "SetupFiles is already copied. Skipping..."
            }

            # Copy Office 2019
            if (-not (Test-Path "$tempDir\Office2019")) {
                Add-Status "Copying Office 2019 files..."
                $officeSource = "D:\SOFTWARE\OFFICE\Office 2019"
                if (Test-Path $officeSource) {
                    Copy-Item -Path "$officeSource\*" -Destination "$tempDir\Office2019" -Recurse -Force
                    Add-Status "Office 2019 has been copied successfully!"
                }
                else {
                    Add-Status "Warning: Office source folder not found at $officeSource"
                }
            }
            else {
                Add-Status "Office 2019 is already copied. Skipping..."
            }

            # Copy Unikey
            if (-not (Test-Path "C:\unikey46RC2-230919-win64")) {
                Add-Status "Copying Unikey files..."
                $unikeySource = "D:\SOFTWARE\PAYOO\unikey46RC2-230919-win64"
                if (Test-Path $unikeySource) {
                    Copy-Item -Path $unikeySource -Destination "C:\unikey46RC2-230919-win64" -Recurse -Force
                    Add-Status "Unikey has been copied successfully!"
                }
                else {
                    Add-Status "Warning: Unikey source folder not found at $unikeySource"
                }
            }
            else {
                Add-Status "Unikey is already copied. Skipping..."
            }

            # Download Microsoft Teams
            $teamsSetupPath = "$env:USERPROFILE\Downloads\TeamsSetup_c_w_.exe"
            if (-not (Test-Path $teamsSetupPath)) {
                Add-Status "Downloading Microsoft Teams from Microsoft website..."
                try {
                    $teamsUrl = "https://go.microsoft.com/fwlink/p/?LinkID=2187327&clcid=0x409&culture=en-us&country=US"
                    Invoke-WebRequest -Uri $teamsUrl -OutFile $teamsSetupPath -UseBasicParsing
                    Add-Status "Microsoft Teams has been downloaded successfully!"
                }
                catch {
                    Add-Status "Warning: Failed to download Microsoft Teams. Error: $_"
                }
            }
            else {
                Add-Status "Microsoft Teams installer already exists. Skipping download..."
            }

            # Copy ForceScout
            $forceScoutDest = "$env:USERPROFILE\Downloads\SC-wKgXWicTb0XhUSNethaFN0vkhji53AY5mektJ7O_RSOdc8bEUVIEAAH_OewU.exe"
            if (-not (Test-Path $forceScoutDest)) {
                Add-Status "Copying ForceScout file..."
                $forceScoutSource = "D:\SOFTWARE\PAYOO\SC-wKgXWicTb0XhUSNethaFN0vkhji53AY5mektJ7O_RSOdc8bEUVIEAAH_OewU.exe"
                if (Test-Path $forceScoutSource) {
                    Copy-Item -Path $forceScoutSource -Destination $forceScoutDest -Force
                    Add-Status "ForceScout has been copied successfully!"
                }
                else {
                    Add-Status "Warning: ForceScout source file not found at $forceScoutSource"
                }
            }
            else {
                Add-Status "ForceScout is already copied. Skipping..."
            }

            # Copy Trellix
            $trellixDest = "$env:USERPROFILE\Downloads\TrellixSmartInstall.exe"
            if (-not (Test-Path $trellixDest)) {
                Add-Status "Copying Trellix file..."
                $trellixSource = "D:\SOFTWARE\PAYOO\TrellixSmartInstall.exe"
                if (Test-Path $trellixSource) {
                    Copy-Item -Path $trellixSource -Destination $trellixDest -Force
                    Add-Status "Trellix has been copied successfully!"
                }
                else {
                    Add-Status "Warning: Trellix source file not found at $trellixSource"
                }
            }
            else {
                Add-Status "Trellix is already copied. Skipping..."
            }

            # Copy MDM
            $mdmDest = "$env:USERPROFILE\Downloads\ManageEngine_MDMLaptopEnrollment"
            if (-not (Test-Path $mdmDest)) {
                Add-Status "Copying MDM files..."
                $mdmSource = "D:\SOFTWARE\PAYOO\ManageEngine_MDMLaptopEnrollment"
                if (Test-Path $mdmSource) {
                    Copy-Item -Path $mdmSource -Destination $mdmDest -Recurse -Force
                    Add-Status "MDM has been copied successfully!"
                }
                else {
                    Add-Status "Warning: MDM source folder not found at $mdmSource"
                }
            }
            else {
                Add-Status "MDM is already copied. Skipping..."
            }

            # Copy device-specific agent
            if ($deviceType -eq "Desktop") {
                $agentDest = "$env:USERPROFILE\Downloads\Desktop Agent.exe"
                if (-not (Test-Path $agentDest)) {
                    Add-Status "Copying Desktop Agent file..."
                    $agentSource = "D:\SOFTWARE\PAYOO\Desktop Agent.exe"
                    if (Test-Path $agentSource) {
                        Copy-Item -Path $agentSource -Destination $agentDest -Force
                        Add-Status "Desktop Agent has been copied successfully!"
                    }
                    else {
                        Add-Status "Warning: Desktop Agent source file not found at $agentSource"
                    }
                }
                else {
                    Add-Status "Desktop Agent is already copied. Skipping..."
                }
            }
            else {
                $agentDest = "$env:USERPROFILE\Downloads\Laptop Agent.exe"
                if (-not (Test-Path $agentDest)) {
                    Add-Status "Copying Laptop Agent file..."
                    $agentSource = "D:\SOFTWARE\PAYOO\Laptop Agent.exe"
                    if (Test-Path $agentSource) {
                        Copy-Item -Path $agentSource -Destination $agentDest -Force
                        Add-Status "Laptop Agent has been copied successfully!"
                    }
                    else {
                        Add-Status "Warning: Laptop Agent source file not found at $agentSource"
                    }
                }
                else {
                    Add-Status "Laptop Agent is already copied. Skipping..."
                }
            }

            Add-Status "All files have been copied successfully."
            return $true
        }
        catch {
            Add-Status "Error during file copy: $_"
            return $false
        }
    }

    # Function to install software
    function Install-Software {
        param (
            [string]$deviceType # "Desktop" or "Laptop"
        )

        try {
            $tempDir = "$env:USERPROFILE\Downloads\SETUP"

            # Install 7-Zip
            if (-not (Test-Path "$env:ProgramFiles\7-Zip\7zFM.exe")) {
                Add-Status "Installing 7-Zip..."
                $zipSetup = "$tempDir\Software\7z2408-x64.exe"
                if (Test-Path $zipSetup) {
                    Start-Process -FilePath $zipSetup -ArgumentList "/S" -Wait
                    Add-Status "7-Zip installed successfully!"
                }
                else {
                    Add-Status "Warning: 7-Zip setup file not found at $zipSetup"
                }
            }
            else {
                Add-Status "7-Zip is already installed. Skipping..."
            }

            # Install Google Chrome
            if (-not (Test-Path "$env:ProgramFiles\Google\Chrome\Application\chrome.exe")) {
                Add-Status "Installing Google Chrome..."
                $chromeSetup = "$tempDir\Software\ChromeSetup.exe"
                if (Test-Path $chromeSetup) {
                    Start-Process -FilePath $chromeSetup -ArgumentList "/silent /install" -Wait
                    Add-Status "Google Chrome installed successfully!"
                }
                else {
                    Add-Status "Warning: Chrome setup file not found at $chromeSetup"
                }
            }
            else {
                Add-Status "Google Chrome is already installed. Skipping..."
            }

            # Install LAPS_x64
            if (-not (Test-Path "$env:ProgramFiles\LAPS\CSE\AdmPwd.dll")) {
                Add-Status "Installing LAPS_x64..."
                $lapsSetup = "$tempDir\Software\LAPS_x64.msi"
                if (Test-Path $lapsSetup) {
                    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$lapsSetup`" /quiet" -Wait
                    Add-Status "LAPS_x64 installed successfully!"
                }
                else {
                    Add-Status "Warning: LAPS setup file not found at $lapsSetup"
                }
            }
            else {
                Add-Status "LAPS_x64 is already installed. Skipping..."
            }

            # Install Foxit Reader
            if (-not (Test-Path "${env:ProgramFiles(x86)}\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe")) {
                Add-Status "Installing Foxit Reader..."
                $foxitSetup = "$tempDir\Software\FoxitPDFReader20243_enu_Setup_Prom.exe"
                if (Test-Path $foxitSetup) {
                    Start-Process -FilePath $foxitSetup -ArgumentList "/silent /install" -Wait
                    Add-Status "Foxit Reader installed successfully!"
                }
                else {
                    Add-Status "Warning: Foxit Reader setup file not found at $foxitSetup"
                }
            }
            else {
                Add-Status "Foxit Reader is already installed. Skipping..."
            }

            # Install Microsoft Office 2019
            if (-not (Test-Path "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE")) {
                Add-Status "Installing Microsoft Office 2019..."
                $officeSetup = "$tempDir\Office2019\setup.exe"
                $officeConfig = "$tempDir\Office2019\configuration.xml"
                if ((Test-Path $officeSetup) -and (Test-Path $officeConfig)) {
                    $currentLocation = Get-Location
                    Set-Location -Path "$tempDir\Office2019"
                    Start-Process -FilePath $officeSetup -ArgumentList "/configure configuration.xml" -Wait
                    Set-Location -Path $currentLocation
                    Add-Status "Microsoft Office 2019 installed successfully!"
                }
                else {
                    Add-Status "Warning: Office setup files not found at $tempDir\Office2019"
                }
            }
            else {
                Add-Status "Microsoft Office 2019 is already installed. Skipping..."
            }

            # Install laptop-specific software
            if ($deviceType -eq "Laptop") {
                # Install Zoom
                if (-not (Test-Path "$env:USERPROFILE\AppData\Roaming\Zoom\bin\Zoom.exe")) {
                    Add-Status "Installing Zoom..."
                    $zoomSetup = "$tempDir\Software\ZoomInstallerFull.exe"
                    if (Test-Path $zoomSetup) {
                        Start-Process -FilePath $zoomSetup -ArgumentList "/silent /install" -Wait
                        Add-Status "Zoom installed successfully!"
                    }
                    else {
                        Add-Status "Warning: Zoom setup file not found at $zoomSetup"
                    }
                }
                else {
                    Add-Status "Zoom is already installed. Skipping..."
                }

                # Install CheckPointVPN
                if (-not (Test-Path "${env:ProgramFiles(x86)}\CheckPoint\Endpoint Connect\TrGUI.exe")) {
                    Add-Status "Installing CheckPointVPN..."
                    $vpnSetup = "$tempDir\Software\CheckPointVPN.msi"
                    if (Test-Path $vpnSetup) {
                        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$vpnSetup`" /quiet" -Wait
                        Add-Status "CheckPointVPN installed successfully!"
                    }
                    else {
                        Add-Status "Warning: CheckPointVPN setup file not found at $vpnSetup"
                    }
                }
                else {
                    Add-Status "CheckPointVPN is already installed. Skipping..."
                }
            }

            Add-Status "All software has been installed successfully."
            return $true
        }
        catch {
            Add-Status "Error during software installation: $_"
            return $false
        }
    }

    function Install-Software-Laptop {
        try {
            Add-Status "DEBUG: Đã vào hàm Install-Software-Laptop"
            $tempDir = "$env:USERPROFILE\Downloads\SETUP"
            # 1. Copy files (như Desktop, chỉ khác Agent)
            Add-Status "DEBUG: Bắt đầu kiểm tra/copy SetupFiles"
            if (-not (Test-Path "$tempDir\Software")) {
                Add-Status "Bắt đầu copy SetupFiles từ D:\SOFTWARE\PAYOO\SETUP đến $tempDir\Software..."
                Copy-Item -Path "D:\SOFTWARE\PAYOO\SETUP" -Destination "$tempDir\Software" -Recurse -Force
                Add-Status "Copy SetupFiles thành công!"
            } else { Add-Status "SetupFiles đã được copy. Bỏ qua..." }
            Add-Status "DEBUG: Bắt đầu kiểm tra/copy Office 2019"
            if (-not (Test-Path "$tempDir\Office2019")) {
                Add-Status "Bắt đầu copy Office 2019 từ D:\SOFTWARE\OFFICE\Office 2019\* đến $tempDir\Office2019..."
                Copy-Item -Path "D:\SOFTWARE\OFFICE\Office 2019\*" -Destination "$tempDir\Office2019" -Recurse -Force
                Add-Status "Copy Office 2019 thành công!"
            } else { Add-Status "Office 2019 đã được copy. Bỏ qua..." }
            Add-Status "DEBUG: Bắt đầu kiểm tra/copy Unikey"
            if (-not (Test-Path "C:\unikey46RC2-230919-win64")) {
                Add-Status "Bắt đầu copy Unikey từ D:\SOFTWARE\PAYOO\unikey46RC2-230919-win64 đến C:\unikey46RC2-230919-win64..."
                Copy-Item -Path "D:\SOFTWARE\PAYOO\unikey46RC2-230919-win64" -Destination "C:\unikey46RC2-230919-win64" -Recurse -Force
                Add-Status "Copy Unikey thành công!"
            } else { Add-Status "Unikey đã được copy. Bỏ qua..." }
            Add-Status "DEBUG: Bắt đầu kiểm tra/copy MSTeamsSetup"
            if (-not (Test-Path "C:\MSTeamsSetup")) {
                Add-Status "Bắt đầu copy MSTeamsSetup từ D:\SOFTWARE\PAYOO\MSTeamsSetup đến C:\MSTeamsSetup..."
                Copy-Item -Path "D:\SOFTWARE\PAYOO\MSTeamsSetup" -Destination "C:\MSTeamsSetup" -Recurse -Force
                Add-Status "Copy MSTeamsSetup thành công!"
            } else { Add-Status "MSTeamsSetup đã được copy. Bỏ qua..." }
            Add-Status "DEBUG: Bắt đầu kiểm tra/copy ForceScout"
            $forceScoutDest = "$env:USERPROFILE\Downloads\SC-wKgXWicTb0XhUSNethaFN0vkhji53AY5mektJ7O_RSOdc8bEUVIEAAH_OewU.exe"
            if (-not (Test-Path $forceScoutDest)) {
                Add-Status "Bắt đầu copy ForceScout từ D:\SOFTWARE\PAYOO\SC-wKgXWicTb0XhUSNethaFN0vkhji53AY5mektJ7O_RSOdc8bEUVIEAAH_OewU.exe đến $forceScoutDest..."
                Copy-Item -Path "D:\SOFTWARE\PAYOO\SC-wKgXWicTb0XhUSNethaFN0vkhji53AY5mektJ7O_RSOdc8bEUVIEAAH_OewU.exe" -Destination $forceScoutDest -Force
                Add-Status "Copy ForceScout thành công!"
            } else { Add-Status "ForceScout đã được copy. Bỏ qua..." }
            Add-Status "DEBUG: Bắt đầu kiểm tra/copy Trellix"
            $trellixDest = "$env:USERPROFILE\Downloads\TrellixSmartInstall.exe"
            if (-not (Test-Path $trellixDest)) {
                Add-Status "Bắt đầu copy Trellix từ D:\SOFTWARE\PAYOO\TrellixSmartInstall.exe đến $trellixDest..."
                Copy-Item -Path "D:\SOFTWARE\PAYOO\TrellixSmartInstall.exe" -Destination $trellixDest -Force
                Add-Status "Copy Trellix thành công!"
            } else { Add-Status "Trellix đã được copy. Bỏ qua..." }
            Add-Status "DEBUG: Bắt đầu kiểm tra/copy MDM"
            $mdmDest = "$env:USERPROFILE\Downloads\ManageEngine_MDMLaptopEnrollment"
            if (-not (Test-Path $mdmDest)) {
                Add-Status "Bắt đầu copy MDM từ D:\SOFTWARE\PAYOO\ManageEngine_MDMLaptopEnrollment đến $mdmDest..."
                Copy-Item -Path "D:\SOFTWARE\PAYOO\ManageEngine_MDMLaptopEnrollment" -Destination $mdmDest -Recurse -Force
                Add-Status "Copy MDM thành công!"
            } else { Add-Status "MDM đã được copy. Bỏ qua..." }
            Add-Status "DEBUG: Bắt đầu kiểm tra/copy Laptop Agent"
            $agentDest = "$env:USERPROFILE\Downloads\Laptop Agent.exe"
            if (-not (Test-Path $agentDest)) {
                Add-Status "Bắt đầu copy Laptop Agent từ D:\SOFTWARE\PAYOO\Laptop Agent.exe đến $agentDest..."
                Copy-Item -Path "D:\SOFTWARE\PAYOO\Laptop Agent.exe" -Destination $agentDest -Force
                Add-Status "Copy Laptop Agent thành công!"
            } else { Add-Status "Laptop Agent đã được copy. Bỏ qua..." }

            # 2. Uninstall OneDrive if present
            Add-Status "DEBUG: Bắt đầu kiểm tra/gỡ OneDrive"
            if (Test-Path "$env:UserProfile\OneDrive") {
                Add-Status "Bắt đầu gỡ cài đặt Microsoft OneDrive..."
                Start-Process -FilePath "$env:SystemRoot\System32\OneDriveSetup.exe" -ArgumentList "/uninstall" -Wait
                Add-Status "Gỡ cài đặt OneDrive thành công!"
            } else { Add-Status "OneDrive không được cài đặt hoặc đã được gỡ. Bỏ qua..." }

            # 3. Install software
            Add-Status "DEBUG: Bắt đầu kiểm tra/cài đặt phần mềm"
            $swDir = "$tempDir\Software"
            # 7-Zip
            if (-not (Test-Path "$env:ProgramFiles\7-Zip\7zFM.exe")) {
                Add-Status "Bắt đầu cài đặt 7-Zip từ $swDir\7z2408-x64.exe..."
                $zipSetup = "$swDir\7z2408-x64.exe"
                if (Test-Path $zipSetup) { Start-Process -FilePath $zipSetup -ArgumentList "/S" -Wait; Add-Status "Cài đặt 7-Zip thành công!" } else { Add-Status "Không tìm thấy file cài đặt 7-Zip!" }
            } else { Add-Status "7-Zip đã được cài đặt. Bỏ qua..." }
            # Chrome
            if (-not (Test-Path "$env:ProgramFiles\Google\Chrome\Application\chrome.exe")) {
                Add-Status "Bắt đầu cài đặt Google Chrome từ $swDir\ChromeSetup.exe..."
                $chromeSetup = "$swDir\ChromeSetup.exe"
                if (Test-Path $chromeSetup) { Start-Process -FilePath $chromeSetup -ArgumentList "/silent /install" -Wait; Add-Status "Cài đặt Google Chrome thành công!" } else { Add-Status "Không tìm thấy file cài đặt Chrome!" }
            } else { Add-Status "Google Chrome đã được cài đặt. Bỏ qua..." }
            # LAPS_x64
            if (-not (Test-Path "$env:ProgramFiles\LAPS\CSE\AdmPwd.dll")) {
                Add-Status "Bắt đầu cài đặt LAPS_x64 từ $swDir\LAPS_x64.msi..."
                $lapsSetup = "$swDir\LAPS_x64.msi"
                if (Test-Path $lapsSetup) { Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$lapsSetup`" /quiet" -Wait; Add-Status "Cài đặt LAPS_x64 thành công!" } else { Add-Status "Không tìm thấy file cài đặt LAPS_x64!" }
            } else { Add-Status "LAPS_x64 đã được cài đặt. Bỏ qua..." }
            # Foxit Reader
            if (-not (Test-Path "${env:ProgramFiles(x86)}\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe")) {
                Add-Status "Bắt đầu cài đặt Foxit Reader từ $swDir\FoxitPDFReader20243_enu_Setup_Prom.exe..."
                $foxitSetup = "$swDir\FoxitPDFReader20243_enu_Setup_Prom.exe"
                if (Test-Path $foxitSetup) { Start-Process -FilePath $foxitSetup -ArgumentList "/silent /install" -Wait; Add-Status "Cài đặt Foxit Reader thành công!" } else { Add-Status "Không tìm thấy file cài đặt Foxit Reader!" }
            } else { Add-Status "Foxit Reader đã được cài đặt. Bỏ qua..." }
            # Office 2019
            if (-not (Test-Path "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE")) {
                Add-Status "Bắt đầu cài đặt Microsoft Office 2019 từ $tempDir\Office2019\setup.exe..."
                $officeSetup = "$tempDir\Office2019\setup.exe"
                $officeConfig = "$tempDir\Office2019\configuration.xml"
                if ((Test-Path $officeSetup) -and (Test-Path $officeConfig)) {
                    $currentLocation = Get-Location
                    Set-Location -Path "$tempDir\Office2019"
                    Start-Process -FilePath $officeSetup -ArgumentList "/configure configuration.xml" -Wait
                    Set-Location -Path $currentLocation
                    Add-Status "Cài đặt Microsoft Office 2019 thành công!"
                } else { Add-Status "Không tìm thấy file cài đặt Office!" }
            } else { Add-Status "Microsoft Office 2019 đã được cài đặt. Bỏ qua..." }
            return $true
        } catch {
            Add-Status "Lỗi khi cài đặt: $_"
            return $false
        }
    }

}

# Power Options and Firewall
$buttonPowerOptions = New-DynamicButton -text "[3] Control Panel" -x 30 -y 260 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Hide the main menu
    Hide-MainMenu
    # Create Power Options and Firewall form
    $powerForm = New-Object System.Windows.Forms.Form
    $powerForm.Text = "Control Panel Management"
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
            [System.Drawing.Color]::FromArgb(0, 0, 0), # Black at top
            [System.Drawing.Color]::FromArgb(0, 40, 0), # Dark green at bottom
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
        )
        $graphics.FillRectangle($brush, $rect)
        $brush.Dispose()
    }

    # Title label with animation
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "CONTROL PANEL MANAGEMENT"
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
            }
            else {
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
    $btnTimeAndPower = New-DynamicButton -text "Set Time/Timezone and Power" -x 50 -y 80 -width 400 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
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

    # When the form is closed, show the main menu again
    $powerForm.Add_FormClosed({
        Show-MainMenu
    })

    # Show the form
    $powerForm.ShowDialog()
}

# Change / Edit Volume
$buttonChangeVolume = New-DynamicButton -text "[4] Change / Edit Volume" -x 30 -y 340 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Hide the main menu
    Hide-MainMenu
    # Create volume management form
    $volumeForm = New-Object System.Windows.Forms.Form
    $volumeForm.Text = "Volume Management"
    $volumeForm.Size = New-Object System.Drawing.Size(800, 600)
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
    $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(800, 40)
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

    # Drive list label
    $driveListLabel = New-Object System.Windows.Forms.Label
    $driveListLabel.Text = "Available Drives:"
    $driveListLabel.Location = New-Object System.Drawing.Point(20, 70)
    $driveListLabel.Size = New-Object System.Drawing.Size(200, 20)
    $driveListLabel.ForeColor = [System.Drawing.Color]::White
    $driveListLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $volumeForm.Controls.Add($driveListLabel)

    # Drive list box
    $driveListBox = New-Object System.Windows.Forms.ListBox
    $driveListBox.Location = New-Object System.Drawing.Point(20, 100)
    $driveListBox.Size = New-Object System.Drawing.Size(760, 150)
    $driveListBox.BackColor = [System.Drawing.Color]::Black
    $driveListBox.ForeColor = [System.Drawing.Color]::Lime
    $driveListBox.Font = New-Object System.Drawing.Font("Consolas", 10)
    $volumeForm.Controls.Add($driveListBox)

    # Content Panel for function buttons
    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Location = New-Object System.Drawing.Point(20, 320)
    $contentPanel.Size = New-Object System.Drawing.Size(760, 230)
    $contentPanel.BackColor = [System.Drawing.Color]::Black
    $contentPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $volumeForm.Controls.Add($contentPanel)

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(20, 560)
    $statusTextBox.Size = New-Object System.Drawing.Size(760, 30)
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
        $drives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } },
        @{Name = 'VolumeName'; Expression = { $_.VolumeName } },
        @{Name = 'Size (GB)'; Expression = { [math]::round($_.Size / 1GB, 0) } },
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

    # Create buttons horizontally
    # Change Drive Letter button
    $btnChangeDriveLetter = New-DynamicButton -text "[1] Change Drive Letter" -x 20 -y 270 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
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
            $drives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } },
            @{Name = 'VolumeName'; Expression = { $_.VolumeName } },
            @{Name = 'Size (GB)'; Expression = { [math]::round($_.Size / 1GB, 0) } },
            @{Name = 'FreeSpace (GB)'; Expression = { [math]::round($_.FreeSpace / 1GB, 0) } }

            foreach ($drive in $drives) {
                $driveInfo = "$($drive.Name) - $($drive.VolumeName) - Size: $($drive.'Size (GB)') GB - Free: $($drive.'FreeSpace (GB)') GB"
                $driveListBox.Items.Add($driveInfo)
            }

            if ($driveListBox.Items.Count -gt 0) {
                $driveListBox.SelectedIndex = 0
            }

            Add-Status "Found $($drives.Count) drives."
        }
        catch {
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
                    $drives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } },
                    @{Name = 'VolumeName'; Expression = { $_.VolumeName } },
                    @{Name = 'Size (GB)'; Expression = { [math]::round($_.Size / 1GB, 0) } },
                    @{Name = 'FreeSpace (GB)'; Expression = { [math]::round($_.FreeSpace / 1GB, 0) } }

                    foreach ($drive in $drives) {
                        $driveInfo = "$($drive.Name) - $($drive.VolumeName) - Size: $($drive.'Size (GB)') GB - Free: $($drive.'FreeSpace (GB)') GB"
                        $driveListBox.Items.Add($driveInfo)
                    }

                    if ($driveListBox.Items.Count -gt 0) {
                        $driveListBox.SelectedIndex = 0
                    }
                }
                else {
                    Add-ChangeStatus "Error changing drive letter. Exit code: $($process.ExitCode)"
                }
            }
            catch {
                Add-ChangeStatus "Error: $_"
            }
            finally {
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
    $btnShrinkVolume = New-DynamicButton -text "[2] Shrink Volume" -x 180 -y 270 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Create Shrink Volume form
        $shrinkForm = New-Object System.Windows.Forms.Form
        $shrinkForm.Text = "Shrink Volume"
        $shrinkForm.Size = New-Object System.Drawing.Size(600, 630)
        $shrinkForm.StartPosition = "CenterScreen"
        $shrinkForm.BackColor = [System.Drawing.Color]::Black
        $shrinkForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $shrinkForm.MaximizeBox = $false
        $shrinkForm.MinimizeBox = $false
        $shrinkForm.Icon = $mainForm.Icon

        # Thêm xử lý phím Esc để đóng form
        $shrinkForm.KeyPreview = $true
        $shrinkForm.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
                $shrinkForm.Close()
            }
        })

        # Title label
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "SHRINK VOLUME AND CREATE NEW PARTITION"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(600, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $shrinkForm.Controls.Add($titleLabel)

        # Thêm hiệu ứng nhấp nháy cho tiêu đề
        $titleTimer = New-Object System.Windows.Forms.Timer
        $titleTimer.Interval = 800
        $titleTimer.Add_Tick({
            if ($titleLabel.ForeColor -eq [System.Drawing.Color]::Lime) {
                $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 200, 0)
            }
            else {
                $titleLabel.ForeColor = [System.Drawing.Color]::Lime
            }
        })
        $titleTimer.Start()

        # Drive list label
        $driveListLabel = New-Object System.Windows.Forms.Label
        $driveListLabel.Text = "Available Drives:"
        $driveListLabel.Location = New-Object System.Drawing.Point(10, 60)
        $driveListLabel.Size = New-Object System.Drawing.Size(200, 20)
        $driveListLabel.ForeColor = [System.Drawing.Color]::White
        $driveListLabel.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
        $shrinkForm.Controls.Add($driveListLabel)

        # Drive list box
        $driveListBox = New-Object System.Windows.Forms.ListBox
        $driveListBox.Location = New-Object System.Drawing.Point(10, 90)
        $driveListBox.Size = New-Object System.Drawing.Size(560, 150)
        $driveListBox.BackColor = [System.Drawing.Color]::Black
        $driveListBox.ForeColor = [System.Drawing.Color]::Lime
        $driveListBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $shrinkForm.Controls.Add($driveListBox)

        # Populate drive list
        Add-Status "Getting list of drives..."
        try {
            $drives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } },
            @{Name = 'VolumeName'; Expression = { $_.VolumeName } },
            @{Name = 'Size (GB)'; Expression = { [math]::round($_.Size / 1GB, 0) } },
            @{Name = 'FreeSpace (GB)'; Expression = { [math]::round($_.FreeSpace / 1GB, 0) } }

            foreach ($drive in $drives) {
                $driveInfo = "$($drive.Name) - $($drive.VolumeName) - Size: $($drive.'Size (GB)') GB - Free: $($drive.'FreeSpace (GB)') GB"
                $driveListBox.Items.Add($driveInfo)
            }

            if ($driveListBox.Items.Count -gt 0) {
                $driveListBox.SelectedIndex = 0
            }

            Add-Status "Found $($drives.Count) drives."
        }
        catch {
            Add-Status "Error getting drive list: $_"
        }

        # Selected drive letter label
        $selectedDriveLabel = New-Object System.Windows.Forms.Label
        $selectedDriveLabel.Text = "Selected Drive Letter:"
        $selectedDriveLabel.Location = New-Object System.Drawing.Point(10, 250)
        $selectedDriveLabel.Size = New-Object System.Drawing.Size(150, 20)
        $selectedDriveLabel.ForeColor = [System.Drawing.Color]::White
        $selectedDriveLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $shrinkForm.Controls.Add($selectedDriveLabel)

        # Selected drive letter textbox
        $selectedDriveTextBox = New-Object System.Windows.Forms.TextBox
        $selectedDriveTextBox.Location = New-Object System.Drawing.Point(180, 250)
        $selectedDriveTextBox.Size = New-Object System.Drawing.Size(50, 25)
        $selectedDriveTextBox.BackColor = [System.Drawing.Color]::Black
        $selectedDriveTextBox.ForeColor = [System.Drawing.Color]::Lime
        $selectedDriveTextBox.Font = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
        $selectedDriveTextBox.MaxLength = 1
        $selectedDriveTextBox.ReadOnly = $true
        $selectedDriveTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
        $shrinkForm.Controls.Add($selectedDriveTextBox)

        # Partition size options group box
        $partitionGroupBox = New-Object System.Windows.Forms.GroupBox
        $partitionGroupBox.Text = "Choose Partition Size"
        $partitionGroupBox.Location = New-Object System.Drawing.Point(10, 280)
        $partitionGroupBox.Size = New-Object System.Drawing.Size(560, 120)
        $partitionGroupBox.ForeColor = [System.Drawing.Color]::Lime
        $partitionGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $shrinkForm.Controls.Add($partitionGroupBox)

            # 80GB radio button
            $radio80GB = New-Object System.Windows.Forms.RadioButton
            $radio80GB.Text = "80GB (recommended for 256GB drives)"
            $radio80GB.Location = New-Object System.Drawing.Point(20, 30)
            $radio80GB.Size = New-Object System.Drawing.Size(350, 20)
            $radio80GB.ForeColor = [System.Drawing.Color]::White
            $radio80GB.Font = New-Object System.Drawing.Font("Arial", 10)
            $radio80GB.Checked = $true
            $partitionGroupBox.Controls.Add($radio80GB)

            # 200GB radio button
            $radio200GB = New-Object System.Windows.Forms.RadioButton
            $radio200GB.Text = "200GB (recommended for 500GB drives)"
            $radio200GB.Location = New-Object System.Drawing.Point(20, 55)
            $radio200GB.Size = New-Object System.Drawing.Size(350, 20)
            $radio200GB.ForeColor = [System.Drawing.Color]::White
            $radio200GB.Font = New-Object System.Drawing.Font("Arial", 10)
            $partitionGroupBox.Controls.Add($radio200GB)

            # 500GB radio button
            $radio500GB = New-Object System.Windows.Forms.RadioButton
            $radio500GB.Text = "500GB (recommended for 1TB+ drives)"
            $radio500GB.Location = New-Object System.Drawing.Point(20, 80)
            $radio500GB.Size = New-Object System.Drawing.Size(350, 20)
            $radio500GB.ForeColor = [System.Drawing.Color]::White
            $radio500GB.Font = New-Object System.Drawing.Font("Arial", 10)
            $partitionGroupBox.Controls.Add($radio500GB)

            # Custom size radio button
            $radioCustom = New-Object System.Windows.Forms.RadioButton
            $radioCustom.Text = "Custom size (MB):"
            $radioCustom.Location = New-Object System.Drawing.Point(380, 30)
            $radioCustom.Size = New-Object System.Drawing.Size(150, 20)
            $radioCustom.ForeColor = [System.Drawing.Color]::White
            $radioCustom.Font = New-Object System.Drawing.Font("Arial", 10)
            $partitionGroupBox.Controls.Add($radioCustom)

            # Custom size textbox
            $customSizeTextBox = New-Object System.Windows.Forms.TextBox
            $customSizeTextBox.Location = New-Object System.Drawing.Point(380, 55)
            $customSizeTextBox.Size = New-Object System.Drawing.Size(150, 25)
            $customSizeTextBox.BackColor = [System.Drawing.Color]::Black
            $customSizeTextBox.ForeColor = [System.Drawing.Color]::Lime
            $customSizeTextBox.Font = New-Object System.Drawing.Font("Consolas", 11)
            $customSizeTextBox.Text = "102400"  # Default to 100GB in MB
            $customSizeTextBox.Enabled = $false
            $partitionGroupBox.Controls.Add($customSizeTextBox)

            # Thêm xử lý sự kiện khi nhấn Enter trong ô custom size
            $customSizeTextBox.Add_KeyDown({
                if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                    $_.SuppressKeyPress = $true  # Ngăn chặn tiếng "beep"
                    $shrinkButton.PerformClick()  # Kích hoạt nút Shrink
                }
            })

            # Enable/disable custom size textbox based on radio selection
            $radioCustom.Add_CheckedChanged({
                $customSizeTextBox.Enabled = $radioCustom.Checked
            })

            # Disable custom size textbox when other options are selected
            $radio80GB.Add_CheckedChanged({
                if ($radio80GB.Checked) {
                    $customSizeTextBox.Enabled = $false
                }
            })

            $radio200GB.Add_CheckedChanged({
                if ($radio200GB.Checked) {
                    $customSizeTextBox.Enabled = $false
                }
            })

            $radio500GB.Add_CheckedChanged({
                if ($radio500GB.Checked) {
                    $customSizeTextBox.Enabled = $false
                }
            })

        # New partition label
        $newLabelLabel = New-Object System.Windows.Forms.Label
        $newLabelLabel.Text = "New Partition Label:"
        $newLabelLabel.Location = New-Object System.Drawing.Point(10, 415)
        $newLabelLabel.Size = New-Object System.Drawing.Size(150, 20)
        $newLabelLabel.ForeColor = [System.Drawing.Color]::White
        $newLabelLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $shrinkForm.Controls.Add($newLabelLabel)

        # New partition label textbox
        $newLabelTextBox = New-Object System.Windows.Forms.TextBox
        $newLabelTextBox.Location = New-Object System.Drawing.Point(180, 410) # Điểm bắt đầu của ô nhập liệu
        $newLabelTextBox.Size = New-Object System.Drawing.Size(250, 25) # Kích thước của ô nhập liệu
        $newLabelTextBox.BackColor = [System.Drawing.Color]::Black
        $newLabelTextBox.ForeColor = [System.Drawing.Color]::Lime
        $newLabelTextBox.Font = New-Object System.Drawing.Font("Consolas", 11)
        $newLabelTextBox.Text = "GAME"

        # Thêm xử lý sự kiện khi nhấn Enter để thực hiện shrink volume
        $newLabelTextBox.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true  # Ngăn chặn tiếng "beep"
                $shrinkButton.PerformClick()  # Kích hoạt nút Shrink
            }
        })

        $shrinkForm.Controls.Add($newLabelTextBox)

        # Status textbox
        $shrinkStatusTextBox = New-Object System.Windows.Forms.TextBox
        $shrinkStatusTextBox.Multiline = $true
        $shrinkStatusTextBox.ScrollBars = "Vertical"
        $shrinkStatusTextBox.Location = New-Object System.Drawing.Point(10, 500)
        $shrinkStatusTextBox.Size = New-Object System.Drawing.Size(560, 80)
        $shrinkStatusTextBox.BackColor = [System.Drawing.Color]::Black
        $shrinkStatusTextBox.ForeColor = [System.Drawing.Color]::Lime
        $shrinkStatusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
        $shrinkStatusTextBox.ReadOnly = $true
        $shrinkStatusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $shrinkStatusTextBox.Text = "Ready to shrink volume..."
        $shrinkForm.Controls.Add($shrinkStatusTextBox)

        # Function to add status message to the shrink form
        function Add-ShrinkStatus {
            param([string]$message)
            $shrinkStatusTextBox.AppendText("$message`r`n")
            $shrinkStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }

        # Update selected drive when drive is selected
        $driveListBox.Add_SelectedIndexChanged({
                if ($driveListBox.SelectedItem) {
                    $selectedDrive = $driveListBox.SelectedItem.ToString()
                    $driveLetter = $selectedDrive.Substring(0, 1)
                    $selectedDriveTextBox.Text = $driveLetter
                }
            })

        # Shrink button
        $shrinkButton = New-DynamicButton -text "Shrink" -x 50 -y 450 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            $driveLetter = $selectedDriveTextBox.Text.Trim().ToUpper()
            $newLabel = $newLabelTextBox.Text.Trim()

            # Validate input
            if ($driveLetter -eq "") {
                Add-ShrinkStatus "Error: Please select a drive."
                return
            }

            if ($newLabel -eq "") {
                Add-ShrinkStatus "Error: Please enter a label for the new partition."
                return
            }

            # Determine partition size
            $sizeMB = 0
            if ($radio80GB.Checked) {
                $sizeMB = 82020
                Add-ShrinkStatus "Selected 80GB partition."
            }
            elseif ($radio200GB.Checked) {
                $sizeMB = 204955
                Add-ShrinkStatus "Selected 200GB partition."
            }
            elseif ($radio500GB.Checked) {
                $sizeMB = 512000
                Add-ShrinkStatus "Selected 500GB partition."
            }
            elseif ($radioCustom.Checked) {
                # Validate custom size input
                $customSize = $customSizeTextBox.Text.Trim()
                if ($customSize -match '^\d+$') {
                    try {
                        $sizeMB = [int]$customSize
                        if ($sizeMB -lt 1024) {
                            Add-ShrinkStatus "Error: Custom size must be at least 1024 MB (1 GB)."
                            return
                        }

                        # Kiểm tra xem ổ đĩa có đủ dung lượng không
                        $selectedDriveInfo = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DeviceID -eq "$($driveLetter):" }
                        if ($selectedDriveInfo) {
                            $freeSpaceMB = [math]::Floor($selectedDriveInfo.FreeSpace / 1MB)
                            if ($sizeMB -gt $freeSpaceMB) {
                                Add-ShrinkStatus "Error: Not enough free space. Available: $freeSpaceMB MB, Requested: $sizeMB MB."
                                return
                            }
                        }

                        Add-ShrinkStatus "Selected custom size: $sizeMB MB."
                    }
                    catch {
                        Add-ShrinkStatus "Error processing custom size: $_"
                        return
                    }
                } else {
                    Add-ShrinkStatus "Error: Custom size must be a valid number."
                    return
                }
            }

            # Create a batch file that will run diskpart and then set the label
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
                    powershell -command "Get-WmiObject Win32_LogicalDisk -Filter \"DeviceID='%driveLetter%:'\" | Select-Object DeviceID, VolumeName, Size, FreeSpace | Format-List" >> shrink_status.txt

                    del diskpart_output.txt
                    del diskpart_script.txt
                    exit /b %errorlevel%
                )

                echo Diskpart completed successfully. >> shrink_status.txt
                echo. >> shrink_status.txt

                echo Getting available drives after operation... >> shrink_status.txt
                powershell -command "Get-WmiObject Win32_LogicalDisk | Select-Object @{Name='Name';Expression={`$_.DeviceID}}, @{Name='VolumeName';Expression={`$_.VolumeName}}, @{Name='Size (GB)';Expression={[math]::round(`$_.Size/1GB, 0)}}, @{Name='FreeSpace (GB)';Expression={[math]::round(`$_.FreeSpace/1GB, 0)}} | Format-Table -AutoSize | Out-String" >> shrink_status.txt
                echo. >> shrink_status.txt

                echo Cleaning up temporary files... >> shrink_status.txt
                del diskpart_output.txt
                del diskpart_script.txt

                echo Operation completed successfully. >> shrink_status.txt
"@
            Set-Content -Path $batchFilePath -Value $batchContent -Force -Encoding ASCII

            Add-ShrinkStatus "Shrinking drive $driveLetter and creating new partition of $sizeMB MB..."
            Add-ShrinkStatus "Processing... Please wait while the operation completes."

            try {
                # Tạo một process để chạy batch file với quyền admin và ẩn cửa sổ cmd
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = "cmd.exe"
                $psi.Arguments = "/c `"$batchFilePath`""
                $psi.UseShellExecute = $true
                $psi.Verb = "runas"
                $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

                # Chạy process
                $batchProcess = [System.Diagnostics.Process]::Start($psi)
                $batchProcess.WaitForExit()

                # Đọc file status và hiển thị trong status box
                if (Test-Path "shrink_status.txt") {
                    $statusContent = Get-Content "shrink_status.txt" -Raw
                    Add-ShrinkStatus "---- Operation Log ----"
                    Add-ShrinkStatus $statusContent
                    Add-ShrinkStatus "---- End of Log ----"
                    Remove-Item "shrink_status.txt" -Force -ErrorAction SilentlyContinue
                }

                # Check if successful
                if ($batchProcess.ExitCode -eq 0) {
                    Add-ShrinkStatus "Operation completed successfully."
                    Add-Status "Shrunk drive $driveLetter and created new partition."

                    # Refresh drive list
                    Add-ShrinkStatus "Refreshing drive list..."
                    Start-Sleep -Seconds 2

                    # Tìm ổ đĩa mới được tạo
                    $newDriveFound = $false
                    $newDriveLetter = ""

                    # Đợi một chút để đảm bảo hệ thống đã cập nhật
                    Start-Sleep -Seconds 2
                    Add-ShrinkStatus "Scanning for new drives..."

                    # Lấy danh sách ổ đĩa hiện tại
                    $currentDrives = Get-WmiObject Win32_LogicalDisk | Select-Object DeviceID, VolumeName

                    # Tìm ổ đĩa có tên "New Volume" hoặc ổ đĩa trống
                    foreach ($drive in $currentDrives) {
                        if ($drive.DeviceID -ne "$($driveLetter):" -and
                            ($drive.VolumeName -eq "New Volume" -or $drive.VolumeName -eq "")) {
                            $newDriveFound = $true
                            $newDriveLetter = $drive.DeviceID.TrimEnd(":")
                            Add-ShrinkStatus "Found new drive: $newDriveLetter"
                            break
                        }
                    }

                    # Lấy thông tin đầy đủ về các ổ đĩa để hiển thị
                    $updatedDrives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } },
                        @{Name = 'VolumeName'; Expression = { $_.VolumeName } },
                        @{Name = 'Size (GB)'; Expression = { [math]::round($_.Size / 1GB, 0) } },
                        @{Name = 'FreeSpace (GB)'; Expression = { [math]::round($_.FreeSpace / 1GB, 0) } }

                    # Nếu tìm thấy ổ đĩa mới, đổi tên nó
                    if ($newDriveFound) {
                        Add-ShrinkStatus "Renaming new drive $newDriveLetter to $newLabel..."

                        # Khởi tạo biến để theo dõi trạng thái đổi tên
                        $renameSuccess = $false

                        try {
                            # Tạo và chạy script PowerShell để đổi tên ổ đĩa
                            $tempScriptPath = [System.IO.Path]::GetTempFileName() + ".ps1"

                            # Tạo nội dung script đơn giản hơn
                            $scriptContent = @"
                                # Đổi tên ổ đĩa $newDriveLetter thành $newLabel
                                "Renaming drive $newDriveLetter to $newLabel..." | Out-File -FilePath "rename_status.txt" -Encoding ASCII

                                # Phương pháp 1: Sử dụng Set-Volume (Windows 8 trở lên)
                                try {
                                    if (Get-Command Set-Volume -ErrorAction SilentlyContinue) {
                                        Set-Volume -DriveLetter $newDriveLetter -NewFileSystemLabel '$newLabel' -ErrorAction SilentlyContinue
                                        "Successfully renamed using Set-Volume" | Out-File -FilePath "rename_status.txt" -Append
                                        exit 0
                                    }
                                } catch {
                                    # Tiếp tục với phương pháp khác
                                }

                                # Phương pháp 2: Sử dụng Win32_Volume
                                `$volume = Get-WmiObject -Query "SELECT * FROM Win32_Volume WHERE DriveLetter='$newDriveLetter`:'" -ErrorAction SilentlyContinue
                                if (`$volume) {
                                    `$volume.Label = '$newLabel'
                                    `$result = `$volume.Put()
                                    if (`$result.ReturnValue -eq 0) {
                                        "Successfully renamed using Win32_Volume" | Out-File -FilePath "rename_status.txt" -Append
                                        exit 0
                                    }
                                }

                                # Phương pháp 3: Sử dụng Win32_LogicalDisk
                                `$logicalDisk = Get-WmiObject -Query "SELECT * FROM Win32_LogicalDisk WHERE DeviceID='$newDriveLetter`:'" -ErrorAction SilentlyContinue
                                if (`$logicalDisk) {
                                    `$logicalDisk.VolumeName = '$newLabel'
                                    `$result = `$logicalDisk.Put()
                                    if (`$result.ReturnValue -eq 0) {
                                        "Successfully renamed using Win32_LogicalDisk" | Out-File -FilePath "rename_status.txt" -Append
                                        exit 0
                                    }
                                }

                                # Phương pháp 4: Sử dụng lệnh label
                                `$labelProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c label $newDriveLetter`:$newLabel" -WindowStyle Hidden -PassThru -Wait
                                if (`$labelProcess.ExitCode -eq 0) {
                                    "Successfully renamed using label command" | Out-File -FilePath "rename_status.txt" -Append
                                    exit 0
                                }

                                "All rename methods failed" | Out-File -FilePath "rename_status.txt" -Append
                                exit 1
"@

                            # Lưu script vào file tạm
                            Set-Content -Path $tempScriptPath -Value $scriptContent -Force

                            # Chạy script với quyền admin và ẩn cửa sổ
                            $psi = New-Object System.Diagnostics.ProcessStartInfo
                            $psi.FileName = "powershell.exe"
                            $psi.Arguments = "-ExecutionPolicy Bypass -File `"$tempScriptPath`" -WindowStyle Hidden"
                            $psi.UseShellExecute = $true
                            $psi.Verb = "runas"
                            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

                            # Chạy process
                            $renameProcess = [System.Diagnostics.Process]::Start($psi)
                            $renameProcess.WaitForExit()

                            # Đọc file log nếu có
                            if (Test-Path "rename_status.txt") {
                                $renameLog = Get-Content "rename_status.txt" -Raw
                                Add-ShrinkStatus "---- Rename Operation Log ----"
                                Add-ShrinkStatus $renameLog
                                Add-ShrinkStatus "---- End of Rename Log ----"
                                Remove-Item "rename_status.txt" -Force -ErrorAction SilentlyContinue
                            }

                            # Kiểm tra kết quả
                            if ($renameProcess.ExitCode -eq 0) {
                                Add-ShrinkStatus "Successfully renamed drive $newDriveLetter to $newLabel."
                                $renameSuccess = $true
                            } else {
                                Add-ShrinkStatus "Failed to rename drive $newDriveLetter. Please rename it manually."
                            }

                            # Xóa file tạm
                            if (Test-Path $tempScriptPath) {
                                Remove-Item $tempScriptPath -Force -ErrorAction SilentlyContinue
                            }
                        }
                        catch {
                            Add-ShrinkStatus "Error renaming drive: $_"
                            Add-ShrinkStatus "Please rename the drive manually."
                        }

                        # Cập nhật lại danh sách ổ đĩa sau khi đổi tên
                        Start-Sleep -Seconds 2
                        $updatedDrives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } },
                            @{Name = 'VolumeName'; Expression = { $_.VolumeName } },
                            @{Name = 'Size (GB)'; Expression = { [math]::round($_.Size / 1GB, 0) } },
                            @{Name = 'FreeSpace (GB)'; Expression = { [math]::round($_.FreeSpace / 1GB, 0) } }
                    }
                    else {
                        Add-ShrinkStatus "Could not find the newly created drive. Please rename it manually."
                    }

                    # Cập nhật giao diện
                    Add-ShrinkStatus "Updating drive list..."

                    # Xóa và cập nhật danh sách ổ đĩa
                    $driveListBox.Items.Clear()
                    foreach ($drive in $updatedDrives) {
                        $driveInfo = "$($drive.Name) - $($drive.VolumeName) - Size: $($drive.'Size (GB)') GB - Free: $($drive.'FreeSpace (GB)') GB"
                        $driveListBox.Items.Add($driveInfo)
                    }

                    # Chọn ổ đĩa mới trong danh sách nếu tìm thấy
                    if ($newDriveFound) {
                        # Tìm ổ đĩa mới trong danh sách
                        for ($i = 0; $i -lt $driveListBox.Items.Count; $i++) {
                            if ($driveListBox.Items[$i].ToString().StartsWith("$($newDriveLetter):")) {
                                $driveListBox.SelectedIndex = $i
                                break
                            }
                        }
                    }
                    # Nếu không tìm thấy ổ đĩa mới hoặc không thể chọn, chọn ổ đĩa đầu tiên
                    if ($driveListBox.SelectedIndex -lt 0 -and $driveListBox.Items.Count -gt 0) {
                        $driveListBox.SelectedIndex = 0
                    }

                    Add-ShrinkStatus "Drive list updated successfully."
                }
                else {
                    Add-ShrinkStatus "Error: The operation failed with exit code $($batchProcess.ExitCode)"
                    Add-ShrinkStatus "Please check the command prompt window for more details."
                }
            }
            catch {
                Add-ShrinkStatus "Error: $_"
            }
            finally {
                # Clean up temp file
                if (Test-Path $batchFilePath) {
                    Remove-Item $batchFilePath -Force -ErrorAction SilentlyContinue
                }
            }
        }
        $shrinkForm.Controls.Add($shrinkButton)

        # Cancel button
        $cancelButton = New-DynamicButton -text "Cancel" -x 320 -y 450 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
            $shrinkForm.Close()
        }
        $shrinkForm.Controls.Add($cancelButton)

        # Show the form
        Add-Status "Opening Shrink Volume dialog..."
        $shrinkForm.ShowDialog()
    }
    $volumeForm.Controls.Add($btnShrinkVolume)

    # Extend Volume button
    $btnExtendVolume = New-DynamicButton -text "[3] Extend Volume" -x 340 -y 270 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Create Merge Volumes form
        $mergeForm = New-Object System.Windows.Forms.Form
        $mergeForm.Text = "Merge Volumes"
        $mergeForm.Size = New-Object System.Drawing.Size(600, 500)
        $mergeForm.StartPosition = "CenterScreen"
        $mergeForm.BackColor = [System.Drawing.Color]::Black
        $mergeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $mergeForm.MaximizeBox = $false
        $mergeForm.MinimizeBox = $false

        # Thêm xử lý phím Esc để đóng form
        $mergeForm.KeyPreview = $true
        $mergeForm.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
                $mergeForm.Close()
            }
        })

        # Title label
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "Merge Volumes"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(600, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $mergeForm.Controls.Add($titleLabel)

        # Drive list label
        $driveListLabel = New-Object System.Windows.Forms.Label
        $driveListLabel.Text = "Available Drives:"
        $driveListLabel.Location = New-Object System.Drawing.Point(20, 60)
        $driveListLabel.Size = New-Object System.Drawing.Size(200, 20)
        $driveListLabel.ForeColor = [System.Drawing.Color]::White
        $driveListLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $mergeForm.Controls.Add($driveListLabel)

        # Drive list box
        $driveListBox = New-Object System.Windows.Forms.ListBox
        $driveListBox.Location = New-Object System.Drawing.Point(20, 90)
        $driveListBox.Size = New-Object System.Drawing.Size(560, 150)
        $driveListBox.BackColor = [System.Drawing.Color]::Black
        $driveListBox.ForeColor = [System.Drawing.Color]::Lime
        $driveListBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $mergeForm.Controls.Add($driveListBox)

        # Populate drive list
        Add-Status "Getting list of drives..."
        try {
            $drives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } },
            @{Name = 'VolumeName'; Expression = { $_.VolumeName } },
            @{Name = 'Size (GB)'; Expression = { [math]::round($_.Size / 1GB, 0) } },
            @{Name = 'FreeSpace (GB)'; Expression = { [math]::round($_.FreeSpace / 1GB, 0) } }

            foreach ($drive in $drives) {
                $driveInfo = "$($drive.Name) - $($drive.VolumeName) - Size: $($drive.'Size (GB)') GB - Free: $($drive.'FreeSpace (GB)') GB"
                $driveListBox.Items.Add($driveInfo)
            }

            if ($driveListBox.Items.Count -gt 0) {
                $driveListBox.SelectedIndex = 0
            }

            Add-Status "Found $($drives.Count) drives."
        }
        catch {
            Add-Status "Error getting drive list: $_"
        }

        # Source drive label
        $sourceDriveLabel = New-Object System.Windows.Forms.Label
        $sourceDriveLabel.Text = "Source Drive (to delete):"
        $sourceDriveLabel.Location = New-Object System.Drawing.Point(20, 250)
        $sourceDriveLabel.Size = New-Object System.Drawing.Size(150, 20)
        $sourceDriveLabel.ForeColor = [System.Drawing.Color]::White
        $sourceDriveLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $mergeForm.Controls.Add($sourceDriveLabel)

        # Source drive textbox
        $sourceDriveTextBox = New-Object System.Windows.Forms.TextBox
        $sourceDriveTextBox.Location = New-Object System.Drawing.Point(180, 250)
        $sourceDriveTextBox.Size = New-Object System.Drawing.Size(50, 20)
        $sourceDriveTextBox.BackColor = [System.Drawing.Color]::Black
        $sourceDriveTextBox.ForeColor = [System.Drawing.Color]::Lime
        $sourceDriveTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $sourceDriveTextBox.MaxLength = 1

        # Thêm xử lý sự kiện khi nhấn Enter
        $sourceDriveTextBox.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true  # Ngăn chặn tiếng "beep"
                $targetDriveTextBox.Focus()  # Chuyển focus đến ô target drive
            }
        })

        $mergeForm.Controls.Add($sourceDriveTextBox)

        # Target drive label
        $targetDriveLabel = New-Object System.Windows.Forms.Label
        $targetDriveLabel.Text = "Target Drive (to expand):"
        $targetDriveLabel.Location = New-Object System.Drawing.Point(20, 280)
        $targetDriveLabel.Size = New-Object System.Drawing.Size(150, 20)
        $targetDriveLabel.ForeColor = [System.Drawing.Color]::White
        $targetDriveLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $mergeForm.Controls.Add($targetDriveLabel)

        # Target drive textbox
        $targetDriveTextBox = New-Object System.Windows.Forms.TextBox
        $targetDriveTextBox.Location = New-Object System.Drawing.Point(180, 280)
        $targetDriveTextBox.Size = New-Object System.Drawing.Size(50, 20)
        $targetDriveTextBox.BackColor = [System.Drawing.Color]::Black
        $targetDriveTextBox.ForeColor = [System.Drawing.Color]::Lime
        $targetDriveTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $targetDriveTextBox.MaxLength = 1

        # Thêm xử lý sự kiện khi nhấn Enter
        $targetDriveTextBox.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true  # Ngăn chặn tiếng "beep"
                $mergeButton.PerformClick()  # Kích hoạt nút Merge
            }
        })

        $mergeForm.Controls.Add($targetDriveTextBox)

        # Warning label
        $warningLabel = New-Object System.Windows.Forms.Label
        $warningLabel.Text = "WARNING: This will DELETE the source drive and all its data!"
        $warningLabel.Location = New-Object System.Drawing.Point(20, 310)
        $warningLabel.Size = New-Object System.Drawing.Size(560, 20)
        $warningLabel.ForeColor = [System.Drawing.Color]::Red
        $warningLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $mergeForm.Controls.Add($warningLabel)

        # Status textbox
        $mergeStatusTextBox = New-Object System.Windows.Forms.TextBox
        $mergeStatusTextBox.Multiline = $true
        $mergeStatusTextBox.ScrollBars = "Vertical"
        $mergeStatusTextBox.Location = New-Object System.Drawing.Point(20, 390)
        $mergeStatusTextBox.Size = New-Object System.Drawing.Size(560, 70)
        $mergeStatusTextBox.BackColor = [System.Drawing.Color]::Black
        $mergeStatusTextBox.ForeColor = [System.Drawing.Color]::Lime
        $mergeStatusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
        $mergeStatusTextBox.ReadOnly = $true
        $mergeStatusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $mergeStatusTextBox.Text = "Ready to merge volumes..."
        $mergeForm.Controls.Add($mergeStatusTextBox)

        # Function to add status message to the merge form
        function Add-MergeStatus {
            param(
                [Parameter(Mandatory=$true)]
                [string]$message,

                [Parameter(Mandatory=$false)]
                [switch]$NoNewLine,

                [Parameter(Mandatory=$false)]
                [switch]$ClearLine
            )

            if ($ClearLine) {
                # Xóa dòng cuối cùng
                $text = $mergeStatusTextBox.Text
                $lastNewLinePos = $text.LastIndexOf("`r`n")
                if ($lastNewLinePos -ge 0) {
                    $mergeStatusTextBox.Text = $text.Substring(0, $lastNewLinePos + 2)
                }
            }

            if ($NoNewLine) {
                $mergeStatusTextBox.AppendText("$message")
            } else {
                $mergeStatusTextBox.AppendText("$message`r`n")
            }

            $mergeStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }

        # Update selected drive when drive is selected
        $driveListBox.Add_SelectedIndexChanged({
                if ($driveListBox.SelectedItem) {
                    $selectedDrive = $driveListBox.SelectedItem.ToString()
                    $driveLetter = $selectedDrive.Substring(0, 1)

                    # If source drive is empty, fill it
                    if ($sourceDriveTextBox.Text -eq "") {
                        $sourceDriveTextBox.Text = $driveLetter
                    }
                    # Otherwise, if target drive is empty and different from source, fill it
                    elseif ($targetDriveTextBox.Text -eq "" -and $driveLetter -ne $sourceDriveTextBox.Text) {
                        $targetDriveTextBox.Text = $driveLetter
                    }
                }
            })

        # Merge button
        $mergeButton = New-DynamicButton -text "Merge and Extend Volume" -x 20 -y 340 -width 250 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            # Lấy và kiểm tra đầu vào
            $sourceDrive = $sourceDriveTextBox.Text.Trim().ToUpper()
            $targetDrive = $targetDriveTextBox.Text.Trim().ToUpper()

            # Hiển thị thông tin đầu vào để debug
            Add-MergeStatus "Source drive entered: '$sourceDrive'"
            Add-MergeStatus "Target drive entered: '$targetDrive'"

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
                        # Lấy thông tin ổ đĩa logic
                        $sourceLogicalDisk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$($sourceDrive):'"
                        $targetLogicalDisk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$($targetDrive):'"

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
        $mergeForm.Controls.Add($mergeButton)

        # Cancel button
        $cancelButton = New-DynamicButton -text "Cancel" -x 290 -y 340 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
            $mergeForm.Close()
        }
        $mergeForm.Controls.Add($cancelButton)

        # Show the form
        Add-Status "Opening Merge Volumes dialog..."
        $mergeForm.ShowDialog()
    }
    $volumeForm.Controls.Add($btnExtendVolume)

    # Rename Volume button
    $btnRenameVolume = New-DynamicButton -text "[4] Rename Volume" -x 500 -y 270 -width 150 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Create Rename Volume form
        $renameVolumeForm = New-Object System.Windows.Forms.Form
        $renameVolumeForm.Text = "Rename Volume"
        $renameVolumeForm.Size = New-Object System.Drawing.Size(500, 500)
        $renameVolumeForm.StartPosition = "CenterScreen"
        $renameVolumeForm.BackColor = [System.Drawing.Color]::Black
        $renameVolumeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $renameVolumeForm.MaximizeBox = $false
        $renameVolumeForm.MinimizeBox = $false

        # Thêm xử lý phím Esc để đóng form
        $renameVolumeForm.KeyPreview = $true
        $renameVolumeForm.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
                $renameVolumeForm.Close()
            }
        })

        # Title label
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "Rename Volume"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(500, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $renameVolumeForm.Controls.Add($titleLabel)

        # Drive list label
        $driveListLabel = New-Object System.Windows.Forms.Label
        $driveListLabel.Text = "Available Drives:"
        $driveListLabel.Location = New-Object System.Drawing.Point(20, 60)
        $driveListLabel.Size = New-Object System.Drawing.Size(200, 20)
        $driveListLabel.ForeColor = [System.Drawing.Color]::White
        $driveListLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $renameVolumeForm.Controls.Add($driveListLabel)

        # Drive list box
        $driveListBox = New-Object System.Windows.Forms.ListBox
        $driveListBox.Location = New-Object System.Drawing.Point(20, 90)
        $driveListBox.Size = New-Object System.Drawing.Size(460, 150)
        $driveListBox.BackColor = [System.Drawing.Color]::Black
        $driveListBox.ForeColor = [System.Drawing.Color]::Lime
        $driveListBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $renameVolumeForm.Controls.Add($driveListBox)

        # Populate drive list
        Add-Status "Getting list of drives..."
        try {
            $drives = Get-WmiObject Win32_LogicalDisk | Select-Object @{Name = 'Name'; Expression = { $_.DeviceID } },
            @{Name = 'VolumeName'; Expression = { $_.VolumeName } },
            @{Name = 'Size (GB)'; Expression = { [math]::round($_.Size / 1GB, 0) } },
            @{Name = 'FreeSpace (GB)'; Expression = { [math]::round($_.FreeSpace / 1GB, 0) } }

            foreach ($drive in $drives) {
                $driveInfo = "$($drive.Name) - $($drive.VolumeName) - Size: $($drive.'Size (GB)') GB - Free: $($drive.'FreeSpace (GB)') GB"
                $driveListBox.Items.Add($driveInfo)
            }

            if ($driveListBox.Items.Count -gt 0) {
                $driveListBox.SelectedIndex = 0
            }

            Add-Status "Found $($drives.Count) drives."
        }
        catch {
            Add-Status "Error getting drive list: $_"
        }

        # Drive letter label
        $driveLetterLabel = New-Object System.Windows.Forms.Label
        $driveLetterLabel.Text = "Drive Letter:"
        $driveLetterLabel.Location = New-Object System.Drawing.Point(20, 250)
        $driveLetterLabel.Size = New-Object System.Drawing.Size(100, 20)
        $driveLetterLabel.ForeColor = [System.Drawing.Color]::White
        $driveLetterLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $renameVolumeForm.Controls.Add($driveLetterLabel)

        # Drive letter textbox
        $driveLetterTextBox = New-Object System.Windows.Forms.TextBox
        $driveLetterTextBox.Location = New-Object System.Drawing.Point(130, 250)
        $driveLetterTextBox.Size = New-Object System.Drawing.Size(50, 20)
        $driveLetterTextBox.BackColor = [System.Drawing.Color]::Black
        $driveLetterTextBox.ForeColor = [System.Drawing.Color]::Lime
        $driveLetterTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $driveLetterTextBox.MaxLength = 1
        $driveLetterTextBox.ReadOnly = $true
        $renameVolumeForm.Controls.Add($driveLetterTextBox)

        # New label label
        $newLabelLabel = New-Object System.Windows.Forms.Label
        $newLabelLabel.Text = "New Label:"
        $newLabelLabel.Location = New-Object System.Drawing.Point(20, 280)
        $newLabelLabel.Size = New-Object System.Drawing.Size(100, 20)
        $newLabelLabel.ForeColor = [System.Drawing.Color]::White
        $newLabelLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $renameVolumeForm.Controls.Add($newLabelLabel)

        # New label textbox
        $newLabelTextBox = New-Object System.Windows.Forms.TextBox
        $newLabelTextBox.Location = New-Object System.Drawing.Point(130, 280)
        $newLabelTextBox.Size = New-Object System.Drawing.Size(350, 20)
        $newLabelTextBox.BackColor = [System.Drawing.Color]::Black
        $newLabelTextBox.ForeColor = [System.Drawing.Color]::Lime
        $newLabelTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)

        # Thêm xử lý sự kiện khi nhấn Enter để rename volume
        $newLabelTextBox.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true  # Ngăn chặn tiếng "beep"
                $renameButton.PerformClick()  # Kích hoạt nút Rename
            }
        })

        $renameVolumeForm.Controls.Add($newLabelTextBox)

        # Status textbox
        $renameStatusTextBox = New-Object System.Windows.Forms.TextBox
        $renameStatusTextBox.Multiline = $true
        $renameStatusTextBox.ScrollBars = "Vertical"
        $renameStatusTextBox.Location = New-Object System.Drawing.Point(20, 360)
        $renameStatusTextBox.Size = New-Object System.Drawing.Size(460, 70)
        $renameStatusTextBox.BackColor = [System.Drawing.Color]::Black
        $renameStatusTextBox.ForeColor = [System.Drawing.Color]::Lime
        $renameStatusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
        $renameStatusTextBox.ReadOnly = $true
        $renameStatusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $renameStatusTextBox.Text = "Ready to rename volume..."
        $renameVolumeForm.Controls.Add($renameStatusTextBox)

        # Function to add status message to the rename form
        function Add-RenameStatus {
            param([string]$message)
            $renameStatusTextBox.AppendText("$message`r`n")
            $renameStatusTextBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }

        # Update selected drive when drive is selected
        $driveListBox.Add_SelectedIndexChanged({
                if ($driveListBox.SelectedItem) {
                    $selectedDrive = $driveListBox.SelectedItem.ToString()
                    $driveLetter = $selectedDrive.Substring(0, 1)
                    $driveLetterTextBox.Text = $driveLetter

                    # Get current volume name
                    $currentVolumeName = ""
                    foreach ($drive in $drives) {
                        if ($drive.Name -eq "$($driveLetter):") {
                            $currentVolumeName = $drive.VolumeName
                            break
                        }
                    }

                    # Set current volume name as default text
                    $newLabelTextBox.Text = $currentVolumeName
                }
            })

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
        $renameButton = New-DynamicButton -text "Rename Volume" -x 20 -y 320 -width 200 -height 30 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            # Lấy thông tin từ form
            $driveLetter = $driveLetterTextBox.Text.Trim().ToUpper()
            # Đảm bảo không có khoảng trắng ở đầu và cuối tên nhãn
            $newLabel = $newLabelTextBox.Text.Trim()

            # Kiểm tra đầu vào
            if ($driveLetter -eq "") {
                Add-RenameStatus "Error: Please select a drive."
                return
            }

            if ($newLabel -eq "") {
                Add-RenameStatus "Error: Please enter a new label."
                return
            }

            # Thông báo bắt đầu đổi tên
            Add-RenameStatus "Renaming drive $driveLetter to $newLabel..."

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
        $renameVolumeForm.Controls.Add($renameButton)

        # Cancel button
        $cancelButton = New-DynamicButton -text "Cancel" -x 240 -y 320 -width 200 -height 30 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
            $renameVolumeForm.Close()
        }
        $renameVolumeForm.Controls.Add($cancelButton)

        # Show the form
        Add-Status "Opening Rename Volume dialog..."
        $renameVolumeForm.ShowDialog()
    }
    $volumeForm.Controls.Add($btnRenameVolume)

    # Return to Main Menu button
    $btnReturn = New-DynamicButton -text "[0] Return to Menu" -x 660 -y 270 -width 120 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
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

# [5] Activate Windows 10 Pro and Office 2019 Pro Plus
$buttonActivate = New-DynamicButton -text "[5] Activate" -x 30 -y 420 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Hide the main menu
    Hide-MainMenu
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
            [System.Drawing.Color]::FromArgb(0, 0, 0), # Black at top
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
            }
            else {
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
        }
        catch {
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
            }
            elseif (Test-Path $office15Path) {
                $officePath = $office15Path
                Add-Status "Found Office15."
            }
            else {
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
        }
        catch {
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
        }
        catch {
            Add-Status "Error upgrading Windows: $_"
        }
    }
    $activateForm.Controls.Add($btnWin10Home)

    # Return to Main Menu button
    $btnReturn = New-DynamicButton -text "[0] Return to Menu" -x 12 -y 220 -width 460 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        $activateForm.Close()
    }
    $activateForm.Controls.Add($btnReturn)

    # When the form is closed, show the main menu again
    $activateForm.Add_FormClosed({
        Show-MainMenu
    })

    # Show the form
    $activateForm.ShowDialog()
}

# [6] Edit Features
$buttonTurnOnFeatures = New-DynamicButton -text "[6] Windows Features" -x 430 -y 100 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Hide the main menu
    Hide-MainMenu
    # Create Windows Features form
    $featuresForm = New-Object System.Windows.Forms.Form
    $featuresForm.Text = "Windows Features"
    $featuresForm.Size = New-Object System.Drawing.Size(500, 550)
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
            [System.Drawing.Color]::FromArgb(0, 0, 0), # Black at top
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
            }
            else {
                $titleLabel.ForeColor = [System.Drawing.Color]::Lime
            }
        })
    $titleTimer.Start()

    $featuresForm.Controls.Add($titleLabel)

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(50, 380)
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

    # Run All Features button
    $btnRunAllFeatures = New-DynamicButton -text "Run All Features" -x 50 -y 80 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 180)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 220)) -pressColor ([System.Drawing.Color]::FromArgb(140, 0, 140)) -clickAction {
        Add-Status "Starting to run all Windows features operations..."

        # 1. Enable .NET Framework 3.5
        Add-Status "Step 1/4: Checking .NET Framework 3.5 status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:NetFx3"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Disabled") {
                Add-Status "Enabling .NET Framework 3.5..."
                $enableCmd = "dism /online /enable-feature /featurename:NetFx3 /all /norestart"
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $enableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -WindowStyle Hidden -Wait
                Add-Status ".NET Framework 3.5 has been enabled."
            }
            else {
                Add-Status ".NET Framework 3.5 is already enabled."
            }
        }
        catch {
            Add-Status "Error enabling .NET Framework 3.5: $_"
        }

        # 2. Enable WCF-HTTP-Activation
        Add-Status "Step 2/4: Checking WCF-HTTP-Activation status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:WCF-HTTP-Activation"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Disabled") {
                Add-Status "Enabling WCF-HTTP-Activation..."
                $enableCmd = "DISM /Online /Enable-Feature /FeatureName:WCF-HTTP-Activation /All /Quiet /NoRestart"
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $enableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -WindowStyle Hidden -Wait
                Add-Status "WCF-HTTP-Activation has been enabled."
            }
            else {
                Add-Status "WCF-HTTP-Activation is already enabled."
            }
        }
        catch {
            Add-Status "Error enabling WCF-HTTP-Activation: $_"
        }

        # 3. Enable WCF-NonHTTP-Activation
        Add-Status "Step 3/4: Checking WCF-NonHTTP-Activation status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:WCF-NonHTTP-Activation"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Disabled") {
                Add-Status "Enabling WCF-NonHTTP-Activation..."
                $enableCmd = "DISM /Online /Enable-Feature /FeatureName:WCF-NonHTTP-Activation /All /Quiet /NoRestart"
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $enableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -WindowStyle Hidden -Wait
                Add-Status "WCF-NonHTTP-Activation has been enabled."
            }
            else {
                Add-Status "WCF-NonHTTP-Activation is already enabled."
            }
        }
        catch {
            Add-Status "Error enabling WCF-NonHTTP-Activation: $_"
        }

        # 4. Disable Internet Explorer 11
        Add-Status "Step 4/4: Checking Internet Explorer 11 status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:Internet-Explorer-Optional-amd64"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Enabled") {
                Add-Status "Disabling Internet Explorer 11..."
                $disableCmd = "dism /online /disable-feature /featurename:Internet-Explorer-Optional-amd64 /norestart"
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $disableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -WindowStyle Hidden -Wait
                Add-Status "Internet Explorer 11 has been disabled."
            }
            else {
                Add-Status "Internet Explorer 11 is already disabled."
            }
        }
        catch {
            Add-Status "Error disabling Internet Explorer 11: $_"
        }

        Add-Status "All Windows features operations completed!"
    }
    $featuresForm.Controls.Add($btnRunAllFeatures)

    # Enable .NET Framework 3.5 button
    $btnEnableNetFx = New-DynamicButton -text "Enable .NET Framework 3.5" -x 50 -y 130 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        Add-Status "Checking .NET Framework 3.5 status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:NetFx3"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Disabled") {
                Add-Status "Enabling .NET Framework 3.5..."
                $enableCmd = "dism /online /enable-feature /featurename:NetFx3 /all /norestart"
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $enableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -WindowStyle Hidden -Wait
                Add-Status ".NET Framework 3.5 has been enabled."
            }
            else {
                Add-Status ".NET Framework 3.5 is already enabled."
            }
        }
        catch {
            Add-Status "Error: $_"
        }
    }
    $featuresForm.Controls.Add($btnEnableNetFx)

    # Enable WCF-HTTP-Activation button
    $btnEnableWcfHttp = New-DynamicButton -text "Enable WCF-HTTP-Activation" -x 50 -y 180 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        Add-Status "Checking WCF-HTTP-Activation status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:WCF-HTTP-Activation"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Disabled") {
                Add-Status "Enabling WCF-HTTP-Activation..."
                $enableCmd = "DISM /Online /Enable-Feature /FeatureName:WCF-HTTP-Activation /All /Quiet /NoRestart"
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $enableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -WindowStyle Hidden -Wait
                Add-Status "WCF-HTTP-Activation has been enabled."
            }
            else {
                Add-Status "WCF-HTTP-Activation is already enabled."
            }
        }
        catch {
            Add-Status "Error: $_"
        }
    }
    $featuresForm.Controls.Add($btnEnableWcfHttp)

    # Enable WCF-NonHTTP-Activation button
    $btnEnableWcfNonHttp = New-DynamicButton -text "Enable WCF-NonHTTP-Activation" -x 50 -y 230 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        Add-Status "Checking WCF-NonHTTP-Activation status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:WCF-NonHTTP-Activation"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Disabled") {
                Add-Status "Enabling WCF-NonHTTP-Activation..."
                $enableCmd = "DISM /Online /Enable-Feature /FeatureName:WCF-NonHTTP-Activation /All /Quiet /NoRestart"
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $enableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -WindowStyle Hidden -Wait
                Add-Status "WCF-NonHTTP-Activation has been enabled."
            }
            else {
                Add-Status "WCF-NonHTTP-Activation is already enabled."
            }
        }
        catch {
            Add-Status "Error: $_"
        }
    }
    $featuresForm.Controls.Add($btnEnableWcfNonHttp)

    # Disable Internet Explorer 11 button
    $btnDisableIE = New-DynamicButton -text "Disable Internet Explorer 11" -x 50 -y 280 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        Add-Status "Checking Internet Explorer 11 status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:Internet-Explorer-Optional-amd64"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Enabled") {
                Add-Status "Disabling Internet Explorer 11..."
                $disableCmd = "dism /online /disable-feature /featurename:Internet-Explorer-Optional-amd64 /norestart"
                Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $disableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -WindowStyle Hidden -Wait
                Add-Status "Internet Explorer 11 has been disabled."
            }
            else {
                Add-Status "Internet Explorer 11 is already disabled."
            }
        }
        catch {
            Add-Status "Error: $_"
        }
    }
    $featuresForm.Controls.Add($btnDisableIE)

    # Return to Main Menu button
    $btnReturn = New-DynamicButton -text "Return to Main Menu" -x 50 -y 330 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        $featuresForm.Close()
    }
    $featuresForm.Controls.Add($btnReturn)

    # When the form is closed, show the main menu again
    $featuresForm.Add_FormClosed({
        Show-MainMenu
    })

    # Show the form
    $featuresForm.ShowDialog()
}

# [7] Rename Device
$buttonRenameDevice = New-DynamicButton -text "[7] Rename Device" -x 430 -y 180 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Hide the main menu
    Hide-MainMenu
    # Create device rename form
    $renameForm = New-Object System.Windows.Forms.Form
    $renameForm.Text = "Rename Device"
    $renameForm.Size = New-Object System.Drawing.Size(500, 480) # Increased height for status box
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

    # Device type selection group box
    $deviceGroupBox = New-Object System.Windows.Forms.GroupBox
    $deviceGroupBox.Text = "Device Type"
    $deviceGroupBox.Font = New-Object System.Drawing.Font("Arial", 10)
    $deviceGroupBox.ForeColor = [System.Drawing.Color]::White
    $deviceGroupBox.Size = New-Object System.Drawing.Size(460, 80)
    $deviceGroupBox.Location = New-Object System.Drawing.Point(20, 110)
    $deviceGroupBox.BackColor = [System.Drawing.Color]::Black

    # Desktop radio button
    $radioDesktop = New-Object System.Windows.Forms.RadioButton
    $radioDesktop.Text = "Desktop (HOD)"
    $radioDesktop.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioDesktop.ForeColor = [System.Drawing.Color]::White
    $radioDesktop.Location = New-Object System.Drawing.Point(20, 30)
    $radioDesktop.Size = New-Object System.Drawing.Size(200, 30)
    $radioDesktop.BackColor = [System.Drawing.Color]::Black
    $radioDesktop.Checked = $true # Default selection

    # Laptop radio button
    $radioLaptop = New-Object System.Windows.Forms.RadioButton
    $radioLaptop.Text = "Laptop (HOL)"
    $radioLaptop.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioLaptop.ForeColor = [System.Drawing.Color]::White
    $radioLaptop.Location = New-Object System.Drawing.Point(240, 30)
    $radioLaptop.Size = New-Object System.Drawing.Size(200, 30)
    $radioLaptop.BackColor = [System.Drawing.Color]::Black

    # Add radio buttons to group box
    $deviceGroupBox.Controls.Add($radioDesktop)
    $deviceGroupBox.Controls.Add($radioLaptop)
    $renameForm.Controls.Add($deviceGroupBox)

    # New name label
    $newNameLabel = New-Object System.Windows.Forms.Label
    $newNameLabel.Text = "New Device Name:"
    $newNameLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $newNameLabel.ForeColor = [System.Drawing.Color]::White
    $newNameLabel.Size = New-Object System.Drawing.Size(150, 30)
    $newNameLabel.Location = New-Object System.Drawing.Point(20, 200)
    $renameForm.Controls.Add($newNameLabel)

    # New name textbox
    $newNameTextBox = New-Object System.Windows.Forms.TextBox
    $newNameTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $newNameTextBox.Size = New-Object System.Drawing.Size(300, 30)
    $newNameTextBox.Location = New-Object System.Drawing.Point(170, 200)
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

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(20, 320)
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
    $renameButton.Location = New-Object System.Drawing.Point(30, 260)
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
    $cancelButton = New-RedButton -text "Cancel" -x 250 -y 260 -width 200 -height 40 -clickAction {
        $renameForm.Close()
    }
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

# [8] Set Password
$buttonSetPassword = New-DynamicButton -text "[8] Set Password" -x 430 -y 260 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Hide the main menu
    Hide-MainMenu
    # Create password setting form
    $passwordForm = New-Object System.Windows.Forms.Form
    $passwordForm.Text = "Set Password"
    $passwordForm.Size = New-Object System.Drawing.Size(500, 340) # Increased height to accommodate the new button
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
    # Password is visible (no UseSystemPasswordChar)
    $passwordTextBox.BackColor = [System.Drawing.Color]::White
    $passwordTextBox.ForeColor = [System.Drawing.Color]::Black
    $passwordForm.Controls.Add($passwordTextBox)

    # Info label for empty password
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = "Leave the password field empty to set a blank password."
    $infoLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $infoLabel.ForeColor = [System.Drawing.Color]::Silver
    $infoLabel.Size = New-Object System.Drawing.Size(450, 20)
    $infoLabel.Location = New-Object System.Drawing.Point(20, 155)
    $passwordForm.Controls.Add($infoLabel)

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
                    # For empty password, use net user command to remove password
                    $command = "net user $currentUser """""
                } else {
                    $command = "net user $currentUser $password"
                }

                # Execute the command
                $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command" -NoNewWindow -Wait -PassThru

                if ($process.ExitCode -eq 0) {
                    # Show success message
                    if ([string]::IsNullOrEmpty($password)) {
                        [System.Windows.Forms.MessageBox]::Show("Password has been removed. User '$currentUser' can now log in without a password.", "Password Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Password has been changed.", "Password Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    }
                    $passwordForm.Close()
                } else {
                    throw "Failed to set password. Exit code: $($process.ExitCode)"
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error setting password: $_`n`nNote: This operation requires administrative privileges.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })
    $passwordForm.Controls.Add($setButton)

    # Cancel button
    $cancelButton = New-RedButton -text "Cancel" -x 250 -y 180 -width 200 -height 40 -clickAction {
        $passwordForm.Close()
    }
    $passwordForm.Controls.Add($cancelButton)

    # Remove Password button
    $removePasswordButton = New-DynamicButton -text "Remove Password" -x 30 -y 240 -width 420 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 100, 150)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 120, 180)) -pressColor ([System.Drawing.Color]::FromArgb(0, 80, 120)) -clickAction {
        try {
            # Confirm with the user
            $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                "Are you sure you want to remove the password for user '$currentUser'?`n`nThis will allow login without a password.",
                "Confirm Password Removal",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)

            if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Command to remove password using net user
                $command = "net user $currentUser """""

                # Execute the command
                $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command" -NoNewWindow -Wait -PassThru

                if ($process.ExitCode -eq 0) {
                    # Show success message
                    [System.Windows.Forms.MessageBox]::Show("Password has been removed. User '$currentUser' can now log in without a password.", "Password Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    $passwordForm.Close()
                } else {
                    throw "Failed to remove password. Exit code: $($process.ExitCode)"
                }
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error removing password: $_`n`nNote: This operation requires administrative privileges.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    $passwordForm.Controls.Add($removePasswordButton)

    # Set the accept button (Enter key)
    $passwordForm.AcceptButton = $setButton
    # Set the cancel button (Escape key)
    $passwordForm.CancelButton = $cancelButton

    # When the form is closed, show the main menu again
    $passwordForm.Add_FormClosed({
        Show-MainMenu
    })

    # Show the form
    $passwordForm.ShowDialog()
}

# [9] Join/Leave Domain
$buttonJoinDomain = New-DynamicButton -text "[9] Domain Management" -x 430 -y 340 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
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

    $radioDomain = New-Object System.Windows.Forms.RadioButton
    $radioDomain.Text = "Join Domain"
    $radioDomain.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioDomain.ForeColor = [System.Drawing.Color]::White
    $radioDomain.Location = New-Object System.Drawing.Point(20, 30)
    $radioDomain.Size = New-Object System.Drawing.Size(120, 30)
    $radioDomain.BackColor = [System.Drawing.Color]::Black
    $radioDomain.Checked = $true

    $radioWorkgroup = New-Object System.Windows.Forms.RadioButton
    $radioWorkgroup.Text = "Join Workgroup"
    $radioWorkgroup.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioWorkgroup.ForeColor = [System.Drawing.Color]::White
    $radioWorkgroup.Location = New-Object System.Drawing.Point(150, 30)
    $radioWorkgroup.Size = New-Object System.Drawing.Size(140, 30)
    $radioWorkgroup.BackColor = [System.Drawing.Color]::Black
    $radioWorkgroup.Checked = $false

    $radioLeaveDomain = New-Object System.Windows.Forms.RadioButton
    $radioLeaveDomain.Text = "Leave Domain"
    $radioLeaveDomain.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioLeaveDomain.ForeColor = [System.Drawing.Color]::White
    $radioLeaveDomain.Location = New-Object System.Drawing.Point(300, 30)
    $radioLeaveDomain.Size = New-Object System.Drawing.Size(140, 30)
    $radioLeaveDomain.BackColor = [System.Drawing.Color]::Black
    $radioLeaveDomain.Checked = $false
    # Only enable Leave Domain option if currently in a domain
    $radioLeaveDomain.Enabled = $isDomain

    $groupBox.Controls.Add($radioDomain)
    $groupBox.Controls.Add($radioWorkgroup)
    $groupBox.Controls.Add($radioLeaveDomain)
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
    $nameTextBox.BackColor = [System.Drawing.Color]::White
    $nameTextBox.ForeColor = [System.Drawing.Color]::Black
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
    $usernameTextBox.BackColor = [System.Drawing.Color]::White
    $usernameTextBox.ForeColor = [System.Drawing.Color]::Black
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
    $passwordTextBox.BackColor = [System.Drawing.Color]::White
    $passwordTextBox.ForeColor = [System.Drawing.Color]::Black
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
                $joinButton.Text = "Join"
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
                $joinButton.Text = "Join"
                $joinButton.Location = New-Object System.Drawing.Point(30, 280)
                $cancelButton.Location = New-Object System.Drawing.Point(250, 280)
                $joinForm.Size = New-Object System.Drawing.Size(500, 380)
            }
        })

    $radioLeaveDomain.Add_CheckedChanged({
            if ($radioLeaveDomain.Checked) {
                $nameLabel.Text = "New Workgroup Name:"
                $nameTextBox.Text = "WORKGROUP"
                $usernameLabel.Visible = $false
                $usernameTextBox.Visible = $false
                $passwordLabel.Visible = $false
                $passwordTextBox.Visible = $false
                $joinButton.Text = "Leave Domain"
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
                elseif ($radioWorkgroup.Checked) {
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
                else {
                    # Leave domain and join workgroup
                    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                        "Are you sure you want to leave the current domain and join the workgroup '$name'? Your computer will restart after this operation.",
                        "Confirm Leave Domain",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Question)

                    if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                        # Command to leave domain and join workgroup
                        $command = "Remove-Computer -WorkgroupName '$name' -Force -Restart"

                        # Create a process to run the command with elevated privileges
                        $psi = New-Object System.Diagnostics.ProcessStartInfo
                        $psi.FileName = "powershell.exe"
                        $psi.Arguments = "-Command Start-Process powershell.exe -ArgumentList '-Command $command' -Verb RunAs"
                        $psi.UseShellExecute = $true
                        $psi.Verb = "runas"

                        # Start the process
                        [System.Diagnostics.Process]::Start($psi)

                        # Show success message
                        [System.Windows.Forms.MessageBox]::Show("Leave domain command has been initiated. If prompted, please allow the elevation request. Your computer will restart to apply the changes.", "Leave Domain", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        $joinForm.Close()
                    }
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error processing domain/workgroup operation: $_`n`nNote: This operation requires administrative privileges.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })
    $joinForm.Controls.Add($joinButton)

    # Cancel button
    $cancelButton = New-DynamicButton -text "Cancel" -x 250 -y 350 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        $joinForm.Close()
    }
    $joinForm.Controls.Add($cancelButton)

    # Set the accept button (Enter key)
    $joinForm.AcceptButton = $joinButton
    # Set the cancel button (Escape key)
    $joinForm.CancelButton = $cancelButton

    # When the form is closed, show the main menu again
    $joinForm.Add_FormClosed({
        Show-MainMenu
    })

    # Show the form
    $joinForm.ShowDialog()
}

# [0] Exit
$buttonExit = New-DynamicButton -text "[0] Exit" -x 430 -y 420 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
    $form.Close()
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