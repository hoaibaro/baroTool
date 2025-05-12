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
$buttonPowerOptions = New-DynamicButton -text "[3] Power Options and Firewall" -x 30 -y 260 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Hide the main menu
    Hide-MainMenu
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
            [System.Drawing.Color]::FromArgb(0, 0, 0), # Black at top
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
            [System.Drawing.Color]::FromArgb(0, 0, 0), # Black at top
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
            }
            else {
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
    $btnShrinkVolume = New-DynamicButton -text "[2] Shrink Volume" -x 50 -y 130 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Create Shrink Volume form
        $shrinkForm = New-Object System.Windows.Forms.Form
        $shrinkForm.Text = "Shrink Volume"
        $shrinkForm.Size = New-Object System.Drawing.Size(600, 600)
        $shrinkForm.StartPosition = "CenterScreen"
        $shrinkForm.BackColor = [System.Drawing.Color]::Black
        $shrinkForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $shrinkForm.MaximizeBox = $false
        $shrinkForm.MinimizeBox = $false

        # Title label
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "Shrink Volume and Create New Partition"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(600, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $shrinkForm.Controls.Add($titleLabel)

        # Drive list label
        $driveListLabel = New-Object System.Windows.Forms.Label
        $driveListLabel.Text = "Available Drives:"
        $driveListLabel.Location = New-Object System.Drawing.Point(20, 60)
        $driveListLabel.Size = New-Object System.Drawing.Size(200, 20)
        $driveListLabel.ForeColor = [System.Drawing.Color]::White
        $driveListLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $shrinkForm.Controls.Add($driveListLabel)

        # Drive list box
        $driveListBox = New-Object System.Windows.Forms.ListBox
        $driveListBox.Location = New-Object System.Drawing.Point(20, 90)
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
        $selectedDriveLabel.Location = New-Object System.Drawing.Point(20, 250)
        $selectedDriveLabel.Size = New-Object System.Drawing.Size(150, 20)
        $selectedDriveLabel.ForeColor = [System.Drawing.Color]::White
        $selectedDriveLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $shrinkForm.Controls.Add($selectedDriveLabel)

        # Selected drive letter textbox
        $selectedDriveTextBox = New-Object System.Windows.Forms.TextBox
        $selectedDriveTextBox.Location = New-Object System.Drawing.Point(180, 250)
        $selectedDriveTextBox.Size = New-Object System.Drawing.Size(50, 20)
        $selectedDriveTextBox.BackColor = [System.Drawing.Color]::Black
        $selectedDriveTextBox.ForeColor = [System.Drawing.Color]::Lime
        $selectedDriveTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $selectedDriveTextBox.MaxLength = 1
        $selectedDriveTextBox.ReadOnly = $true
        $shrinkForm.Controls.Add($selectedDriveTextBox)

        # Partition size options group box
        $partitionGroupBox = New-Object System.Windows.Forms.GroupBox
        $partitionGroupBox.Text = "Choose Partition Size"
        $partitionGroupBox.Location = New-Object System.Drawing.Point(20, 280)
        $partitionGroupBox.Size = New-Object System.Drawing.Size(560, 120)
        $partitionGroupBox.ForeColor = [System.Drawing.Color]::White
        $partitionGroupBox.Font = New-Object System.Drawing.Font("Arial", 10)
        $shrinkForm.Controls.Add($partitionGroupBox)

        # 80GB radio button
        $radio80GB = New-Object System.Windows.Forms.RadioButton
        $radio80GB.Text = "80GB (recommended for 256GB drives)"
        $radio80GB.Location = New-Object System.Drawing.Point(20, 30)
        $radio80GB.Size = New-Object System.Drawing.Size(300, 20)
        $radio80GB.ForeColor = [System.Drawing.Color]::White
        $radio80GB.Checked = $true
        $partitionGroupBox.Controls.Add($radio80GB)

        # 200GB radio button
        $radio200GB = New-Object System.Windows.Forms.RadioButton
        $radio200GB.Text = "200GB (recommended for 500GB drives)"
        $radio200GB.Location = New-Object System.Drawing.Point(20, 55)
        $radio200GB.Size = New-Object System.Drawing.Size(300, 20)
        $radio200GB.ForeColor = [System.Drawing.Color]::White
        $partitionGroupBox.Controls.Add($radio200GB)

        # 500GB radio button
        $radio500GB = New-Object System.Windows.Forms.RadioButton
        $radio500GB.Text = "500GB (recommended for 1TB+ drives)"
        $radio500GB.Location = New-Object System.Drawing.Point(20, 80)
        $radio500GB.Size = New-Object System.Drawing.Size(300, 20)
        $radio500GB.ForeColor = [System.Drawing.Color]::White
        $partitionGroupBox.Controls.Add($radio500GB)

        # New partition label
        $newLabelLabel = New-Object System.Windows.Forms.Label
        $newLabelLabel.Text = "New Partition Label:"
        $newLabelLabel.Location = New-Object System.Drawing.Point(20, 410)
        $newLabelLabel.Size = New-Object System.Drawing.Size(150, 20)
        $newLabelLabel.ForeColor = [System.Drawing.Color]::White
        $newLabelLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $shrinkForm.Controls.Add($newLabelLabel)

        # New partition label textbox
        $newLabelTextBox = New-Object System.Windows.Forms.TextBox
        $newLabelTextBox.Location = New-Object System.Drawing.Point(180, 410)
        $newLabelTextBox.Size = New-Object System.Drawing.Size(200, 20)
        $newLabelTextBox.BackColor = [System.Drawing.Color]::Black
        $newLabelTextBox.ForeColor = [System.Drawing.Color]::Lime
        $newLabelTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
        $newLabelTextBox.Text = "New Volume"
        $shrinkForm.Controls.Add($newLabelTextBox)

        # Status textbox
        $shrinkStatusTextBox = New-Object System.Windows.Forms.TextBox
        $shrinkStatusTextBox.Multiline = $true
        $shrinkStatusTextBox.ScrollBars = "Vertical"
        $shrinkStatusTextBox.Location = New-Object System.Drawing.Point(20, 490)
        $shrinkStatusTextBox.Size = New-Object System.Drawing.Size(560, 70)
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
        $shrinkButton = New-DynamicButton -text "Shrink and Create Partition" -x 20 -y 450 -width 250 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
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

            # Create a batch file that will run diskpart and then set the label
            $batchFilePath = "shrink_volume.bat"

            $batchContent = @"
@echo off
echo ============================================================
echo                  Shrinking Volume $driveLetter
echo ============================================================
echo.

echo Creating diskpart script...
(
    echo select volume $driveLetter
    echo shrink desired=$sizeMB
    echo create partition primary
    echo format fs=ntfs quick
    echo assign
) > diskpart_script.txt

echo Running diskpart...
diskpart /s diskpart_script.txt
if %errorlevel% neq 0 (
    echo Error: Diskpart failed with exit code %errorlevel%
    echo This could be due to insufficient free space or the drive being in use.
    echo Try defragmenting the drive first or closing any applications using the drive.
    pause
    exit /b %errorlevel%
)

echo Diskpart completed successfully.
echo.
echo ============================================================
echo                  Available Drives
echo ============================================================
powershell -command "Get-WmiObject Win32_LogicalDisk | Select-Object @{Name='Name';Expression={`$_.DeviceID}}, @{Name='VolumeName';Expression={`$_.VolumeName}}, @{Name='Size (GB)';Expression={[math]::round(`$_.Size/1GB, 0)}}, @{Name='FreeSpace (GB)';Expression={[math]::round(`$_.FreeSpace/1GB, 0)}} | Format-Table -AutoSize"
echo ============================================================
echo.

set /p drive_letter=Enter the drive letter of the new partition (e.g., D):
set /p new_label=Enter the label for the new partition:

echo Setting label for drive %drive_letter%: to "%new_label%"...
label %drive_letter%: "%new_label%"
if %errorlevel% neq 0 (
    echo Error: Failed to set label.
    pause
    exit /b %errorlevel%
)

echo Successfully set label for drive %drive_letter%: to "%new_label%".
echo Cleaning up temporary files...
del diskpart_script.txt

echo Operation completed successfully.
pause
"@
            Set-Content -Path $batchFilePath -Value $batchContent -Force -Encoding ASCII

            Add-ShrinkStatus "Shrinking drive $driveLetter and creating new partition of $sizeMB MB..."
            Add-ShrinkStatus "A command prompt window will open to complete the operation."
            Add-ShrinkStatus "Please follow the instructions in the command prompt window."

            try {
                # Run the batch file with elevated privileges
                $batchProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$batchFilePath`"" -Verb RunAs -PassThru -Wait

                # Check if successful
                if ($batchProcess.ExitCode -eq 0) {
                    Add-ShrinkStatus "Operation completed successfully."
                    Add-Status "Shrunk drive $driveLetter and created new partition."

                    # Refresh drive list
                    Add-ShrinkStatus "Refreshing drive list..."
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
        $cancelButton = New-DynamicButton -text "Cancel" -x 290 -y 450 -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
            $shrinkForm.Close()
        }
        $shrinkForm.Controls.Add($cancelButton)

        # Show the form
        Add-Status "Opening Shrink Volume dialog..."
        $shrinkForm.ShowDialog()
    }
    $volumeForm.Controls.Add($btnShrinkVolume)

    # Extend Volume button
    $btnExtendVolume = New-DynamicButton -text "[3] Extend Volume" -x 50 -y 180 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Create Merge Volumes form
        $mergeForm = New-Object System.Windows.Forms.Form
        $mergeForm.Text = "Merge Volumes"
        $mergeForm.Size = New-Object System.Drawing.Size(600, 500)
        $mergeForm.StartPosition = "CenterScreen"
        $mergeForm.BackColor = [System.Drawing.Color]::Black
        $mergeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $mergeForm.MaximizeBox = $false
        $mergeForm.MinimizeBox = $false

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
            param([string]$message)
            $mergeStatusTextBox.AppendText("$message`r`n")
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
        $mergeButton = New-DynamicButton -text "Merge Volumes" -x 20 -y 340 -width 250 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            $sourceDrive = $sourceDriveTextBox.Text.Trim().ToUpper()
            $targetDrive = $targetDriveTextBox.Text.Trim().ToUpper()

            # Validate input
            if ($sourceDrive -eq "") {
                Add-MergeStatus "Error: Please enter a source drive letter."
                return
            }

            if ($targetDrive -eq "") {
                Add-MergeStatus "Error: Please enter a target drive letter."
                return
            }

            if ($sourceDrive -eq $targetDrive) {
                Add-MergeStatus "Error: Source and target drives cannot be the same."
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
                Add-MergeStatus "Operation cancelled by user."
                return
            }

            # Create a batch file that will run diskpart
            $batchFilePath = "merge_volumes.bat"

            $batchContent = @"
@echo off
echo ============================================================
echo                  Merging Volumes
echo ============================================================
echo.

echo Checking if drives exist...
powershell -command "(Get-WmiObject Win32_Volume).DriveLetter -contains '%sourceDrive%:\'" >nul 2>&1
if errorlevel 1 (
    echo ERROR: Source drive %sourceDrive% does not exist. Exiting...
    pause
    exit /b 1
)

powershell -command "(Get-WmiObject Win32_Volume).DriveLetter -contains '%targetDrive%:\'" >nul 2>&1
if errorlevel 1 (
    echo ERROR: Target drive %targetDrive% does not exist. Exiting...
    pause
    exit /b 1
)

echo Deleting source drive %sourceDrive%...
(
    echo list volume
    echo select volume %sourceDrive%
    echo detail volume
    echo delete volume override
) > diskpart_script.txt
diskpart /s diskpart_script.txt
if errorlevel 1 (
    echo ERROR: Failed to delete source drive %sourceDrive%.
    echo This could be because:
    echo 1. The volume is in use by Windows or another program
    echo 2. The volume is a system, boot, or recovery volume
    echo 3. The volume has open files or folders
    del diskpart_script.txt
    pause
    exit /b 1
)
del diskpart_script.txt

echo Extending target drive %targetDrive%...
(
    echo list volume
    echo select volume %targetDrive%
    echo detail volume
    echo list disk
    echo select volume %targetDrive%
    echo extend
) > diskpart_script.txt
diskpart /s diskpart_script.txt
if errorlevel 1 (
    echo ERROR: Failed to extend target drive %targetDrive%.
    echo This could be because:
    echo 1. The volumes are not on the same physical disk
    echo 2. There is no unallocated space adjacent to the target volume
    echo 3. The target volume is not a basic volume
    del diskpart_script.txt
    pause
    exit /b 1
)
del diskpart_script.txt

echo Merge completed successfully!
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

            Add-MergeStatus "Merging volumes: deleting drive $sourceDrive and extending drive $targetDrive..."
            Add-MergeStatus "A command prompt window will open to complete the operation."
            Add-MergeStatus "Please follow the instructions in the command prompt window."

            try {
                # Run the batch file with elevated privileges
                $batchProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$batchFilePath`"" -Verb RunAs -PassThru -Wait

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
                    Add-MergeStatus "Error: The operation failed with exit code $($batchProcess.ExitCode)"
                    Add-MergeStatus "Please check the command prompt window for more details."
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
    $btnRenameVolume = New-DynamicButton -text "[4] Rename" -x 50 -y 230 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
        # Create Rename Volume form
        $renameVolumeForm = New-Object System.Windows.Forms.Form
        $renameVolumeForm.Text = "Rename Volume"
        $renameVolumeForm.Size = New-Object System.Drawing.Size(500, 450)
        $renameVolumeForm.StartPosition = "CenterScreen"
        $renameVolumeForm.BackColor = [System.Drawing.Color]::Black
        $renameVolumeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $renameVolumeForm.MaximizeBox = $false
        $renameVolumeForm.MinimizeBox = $false

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
                # Đặt tên nhãn trực tiếp, không qua biến trung gian
                $volume.Label = "$NewLabel"
                $result = $volume.Put()
                if ($result.ReturnValue -eq 0) {
                    return $true
                }
            }

            # Phương pháp 2: Sử dụng Win32_LogicalDisk
            $logicalDisk = Get-WmiObject -Query "SELECT * FROM Win32_LogicalDisk WHERE DeviceID='$DriveLetter`:'"
            if ($logicalDisk) {
                # Đặt tên nhãn trực tiếp, không qua biến trung gian
                $logicalDisk.VolumeName = "$NewLabel"
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

            # Sử dụng chuỗi nối để đảm bảo không có khoảng trắng
            $command = "label " + $DriveLetter + ":" + $NewLabel
            $labelProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command" -WindowStyle Hidden -PassThru -Wait
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
@REM Sử dụng cú pháp đặc biệt để đảm bảo không có khoảng trắng
set drive_letter=$DriveLetter
set new_label=$NewLabel
cmd /c "label %drive_letter%:%new_label%"
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
            $newLabel = $newLabelTextBox.Text

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
    $btnReturn = New-DynamicButton -text "[0] Return to Menu" -x 50 -y 280 -width 400 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
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
$buttonSetPassword = New-DynamicButton -text "[6] Set Password" -x 430 -y 260 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
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
                    # For empty password, use wmic command to remove password requirement
                    $command = "wmic useraccount where name='$currentUser' set passwordrequired=false"
                } else {
                    $command = "net user $currentUser $password"
                }

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
                if ([string]::IsNullOrEmpty($password)) {
                    [System.Windows.Forms.MessageBox]::Show("Password requirement has been removed. User '$currentUser' can now log in without a password. If prompted, please allow the elevation request.", "Password Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Password has been changed. If prompted, please allow the elevation request.", "Password Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
                $passwordForm.Close()
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
                # Command to remove password using net user with an alternative approach
                $command = "wmic useraccount where name='$currentUser' set passwordrequired=false"

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
                [System.Windows.Forms.MessageBox]::Show("Password has been removed. User '$currentUser' can now log in without a password.", "Password Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $passwordForm.Close()
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
$buttonJoinDomain = New-DynamicButton -text "[8] Domain Management" -x 430 -y 340 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
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