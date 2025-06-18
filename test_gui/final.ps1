# ==============================================================================
# BAROPROVIP - VOLUME MANAGEMENT TOOL (FULLY ORGANIZED VERSION)
# ==============================================================================

# SECTION I: ADMIN PRIVILEGES CHECK & INITIALIZATION
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

# SECTION II: UTILITY FUNCTIONS - Các hàm tiện ích chung
# Functions to hide/show the main menu
function Hide-MainMenu {
    $script:form.Hide()
}

function Show-MainMenu {
    $script:form.Show()
    $script:form.BringToFront()
}

# SECTION III: UI CREATION FUNCTIONS - Các hàm tạo giao diện
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

# SECTION IV: MAIN APPLICATION - Ứng dụng chính
# Create main form
$script:form = New-Object System.Windows.Forms.Form
$script:form.Text = "BAOPROVIP - SYSTEM MANAGEMENT"
$script:form.Size = New-Object System.Drawing.Size(850, 550)
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

#================================================
# SECTION 1: RUN ALL OPERATIONS FUNCTIONS - Các hàm thực hiện tất cả các thao tác
# Function to handle Run All operations

# [1] Run All button
$buttonRunAll = New-DynamicButton -text "[1] Run All" -x 30 -y 100 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 250, 0)) -clickAction {
    Invoke-RunAllOperations -mainForm $script:form
}

#================================================
# SECTION 2: INSTALL SOFTWARE DIALOG FUNCTIONS (DEFINED BEFORE USE)
# Function to show Install Software dialog
function Show-InstallSoftwareDialog {
    # Hide the main menu
    Hide-MainMenu
    # Create device type selection form
    $deviceTypeForm = New-Object System.Windows.Forms.Form
    $deviceTypeForm.Text = "Select Device Type"
    $deviceTypeForm.Size = New-Object System.Drawing.Size(485, 480)
    $deviceTypeForm.StartPosition = "CenterScreen"
    $deviceTypeForm.BackColor = [System.Drawing.Color]::Black
    $deviceTypeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $deviceTypeForm.MaximizeBox = $false
    $deviceTypeForm.MinimizeBox = $false

    # Add gradient background
    $deviceTypeForm.Paint = {
        $graphics = $_.Graphics
        $rect = New-Object System.Drawing.Rectangle(0, 0, $deviceTypeForm.Width, $deviceTypeForm.Height)
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $rect,
            [System.Drawing.Color]::FromArgb(0, 0, 0),
            [System.Drawing.Color]::FromArgb(0, 40, 0),
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
        )
        $graphics.FillRectangle($brush, $rect)
        $brush.Dispose()
    }

    # Title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "SELECT DEVICE TYPE"
    $titleLabel.Location = New-Object System.Drawing.Point(110, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(250, 40)
    $titleLabel.ForeColor = [System.Drawing.Color]::Lime
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $deviceTypeForm.Controls.Add($titleLabel)

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(10, 130)
    $statusTextBox.Size = New-Object System.Drawing.Size(450, 300)
    $statusTextBox.BackColor = [System.Drawing.Color]::Black
    $statusTextBox.ForeColor = [System.Drawing.Color]::Lime
    $statusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $statusTextBox.ReadOnly = $true
    $statusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $statusTextBox.Text = "Please select a device type..."
    $deviceTypeForm.Controls.Add($statusTextBox)

    # Function to add status message (DEFINED INSIDE THE DIALOG FUNCTION)
    function Add-Status {
        param([string]$message)

        if ($statusTextBox.Text -eq "Please select a device type...") {
            $statusTextBox.Clear()
        }

        $timestamp = Get-Date -Format "HH:mm:ss"
        $statusTextBox.AppendText("[$timestamp] $message`r`n")
        $statusTextBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Copy-SoftwareFiles function (DEFINED INSIDE THE DIALOG FUNCTION)
    function Copy-SoftwareFiles {
        param ([string]$deviceType)

        try {       
            $tempDir = "$env:USERPROFILE\Downloads\SETUP"
             
            if (-not (Test-Path $tempDir)) {
                Add-Status "Creating temporary folder..."
                New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
                Add-Status "Temporary folder created successfully!"
            }
            else {
                Add-Status "Temporary folder already exists. Skipping..."
            }

            # Check D: drive
            if (-not (Test-Path "D:\")) {
                Add-Status "WARNING: D drive not found. Creating mock installation..."
                
                if (-not (Test-Path "$tempDir\Software")) {
                    New-Item -Path "$tempDir\Software" -ItemType Directory -Force | Out-Null
                    Add-Status "Created mock Software directory"
                }
                 
                if (-not (Test-Path "$tempDir\Office2019")) {
                    New-Item -Path "$tempDir\Office2019" -ItemType Directory -Force | Out-Null
                    Add-Status "Created mock Office2019 directory"
                }
                 
                Add-Status "Copy-SoftwareFiles completed (mock mode)"
                return $true
            }
             
            # Copy SETUP folder from D:\SOFTWARE\PAYOO\SETUP
            if (-not (Test-Path "$tempDir\Software")) {
                $setupSource = "D:\SOFTWARE\PAYOO\SETUP"
                if (Test-Path $setupSource) {
                    Add-Status "Copying setup files from $setupSource..."
                    try {
                        Copy-Item -Path $setupSource -Destination "$tempDir\Software" -Recurse -Force -ErrorAction Stop
                        Add-Status "SetupFiles    has been copied successfully!"
                    }
                    catch {
                        Add-Status "Error copying setup files: $_"
                    }
                }
                else {
                    Add-Status "Warning: Setup source folder not found at $setupSource"
                }
            }
            else {
                Add-Status "SetupFiles    is already copied. Skipping..."
            }

            # Copy Office 2019
            if (-not (Test-Path "$tempDir\Office2019")) {
                $officeSource = "D:\SOFTWARE\OFFICE\Office 2019"
                if (Test-Path $officeSource) {
                    Add-Status "Copying Office 2019 files from $officeSource..."
                    try {
                        New-Item -Path "$tempDir\Office2019" -ItemType Directory -Force | Out-Null
                        Copy-Item -Path "$officeSource\*" -Destination "$tempDir\Office2019" -Recurse -Force -ErrorAction Stop
                        Add-Status "Office 2019   has been copied successfully!"
                    }
                    catch {
                        Add-Status "Error copying Office 2019: $_"
                    }
                }
                else {
                    Add-Status "Warning: Office source folder not found at $officeSource"
                }
            }
            else {
                Add-Status "Office 2019   is already copied. Skipping..."
            }

            # Copy Unikey to C:\ drive
            if (-not (Test-Path "C:\unikey46RC2-230919-win64")) {
                $unikeySource = "D:\SOFTWARE\PAYOO\unikey46RC2-230919-win64"
                if (Test-Path $unikeySource) {
                    Add-Status "Copying Unikey files to C:\ drive..."
                    try {
                        Copy-Item -Path $unikeySource -Destination "C:\unikey46RC2-230919-win64" -Recurse -Force -ErrorAction Stop
                        Add-Status "Unikey        has been copied successfully!"
                    }
                    catch {
                        Add-Status "Error copying Unikey: $_"
                    }
                }
                else {
                    Add-Status "Warning: Unikey source folder not found at $unikeySource"
                }
            }
            else {
                Add-Status "Unikey        is already copied. Skipping..."
            }

            # Copy MSTeamsSetup to C:\ drive
            if (-not (Test-Path "C:\MSTeamsSetup.exe")) {
                $teamsSource = "D:\SOFTWARE\PAYOO\MSTeamsSetup.exe"
                if (Test-Path $teamsSource) {
                    Add-Status "Copying MSTeamsSetup file to C:\ drive..."
                    try {
                        Copy-Item -Path $teamsSource -Destination "C:\MSTeamsSetup.exe" -Force -ErrorAction Stop
                        Add-Status "MSTeamsSetup  has been copied successfully!"
                    }
                    catch {
                        Add-Status "Error copying MSTeamsSetup: $_"
                    }
                }
                else {
                    Add-Status "Warning: MSTeamsSetup source file not found at $teamsSource"
                }
            }
            else {
                Add-Status "MSTeamsSetup  is already copied. Skipping..."
            }

            # Copy ForceScout
            $forceScoutDest = "$env:USERPROFILE\Downloads\ForceScout.exe"
            if (-not (Test-Path $forceScoutDest)) {
                $forceScoutSource = "D:\SOFTWARE\PAYOO\ForceScout.exe"
                if (Test-Path $forceScoutSource) {
                    Add-Status "Copying ForceScout file..."
                    try {
                        Copy-Item -Path $forceScoutSource -Destination $forceScoutDest -Force -ErrorAction Stop
                        Add-Status "ForceScout    has been copied successfully!"
                    }
                    catch {
                        Add-Status "Error copying ForceScout: $_"
                    }
                }
                else {
                    Add-Status "Warning: ForceScout source file not found at $forceScoutSource"
                }
            }
            else {
                Add-Status "ForceScout    is already copied. Skipping..."
            }

            # Copy FalconSensor folder
            $falconDest = "$env:USERPROFILE\Downloads\FalconSensor_Windows_installer (All AV)"
            if (-not (Test-Path $falconDest)) {
                $falconSource = "D:\SOFTWARE\PAYOO\FalconSensor_Windows_installer (All AV)"
                if (Test-Path $falconSource) {
                    Add-Status "Copying FalconSensor folder..."
                    try {
                        Copy-Item -Path $falconSource -Destination $falconDest -Recurse -Force -ErrorAction Stop
                        Add-Status "FalconSensor  has been copied successfully!"
                    }
                    catch {
                        Add-Status "Error copying FalconSensor: $_"
                    }
                }
                else {
                    Add-Status "Warning: FalconSensor source folder not found at $falconSource"
                }
            }
            else {
                Add-Status "FalconSensor  is already copied. Skipping..."
            }

            # Copy device-specific agent
            if ($deviceType -eq "Desktop") {
                $agentDest = "$env:USERPROFILE\Downloads\Desktop Agent.exe"
                if (-not (Test-Path $agentDest)) {
                    $agentSource = "D:\SOFTWARE\PAYOO\Desktop Agent.exe"
                    if (Test-Path $agentSource) {
                        Add-Status "Copying Desktop Agent file..."
                        try {
                            Copy-Item -Path $agentSource -Destination $agentDest -Force -ErrorAction Stop
                            Add-Status "Desktop Agent has been copied successfully!"
                        }
                        catch {
                            Add-Status "Error copying Desktop Agent: $_"
                        }
                    }
                    else {
                        Add-Status "Warning: Desktop Agent source file not found at $agentSource"
                    }
                }
                else {
                    Add-Status "Desktop Agent is already copied. Skipping..."
                }
            }
            elseif ($deviceType -eq "Laptop") {
                # Copy Laptop Agent
                $agentDest = "$env:USERPROFILE\Downloads\Laptop Agent.exe"
                if (-not (Test-Path $agentDest)) {
                    $agentSource = "D:\SOFTWARE\PAYOO\Laptop Agent.exe"
                    if (Test-Path $agentSource) {
                        Add-Status "Copying Laptop Agent file..."
                        try {
                            Copy-Item -Path $agentSource -Destination $agentDest -Force -ErrorAction Stop
                            Add-Status "Laptop Agent  has been copied successfully!"
                        }
                        catch {
                            Add-Status "Error copying Laptop Agent: $_"
                        }
                    }
                    else {
                        Add-Status "Warning: Laptop Agent source file not found at $agentSource"
                    }
                }
                else {
                    Add-Status "Laptop Agent  is already copied. Skipping..."
                }

                # Copy MDM for laptops
                $mdmDest = "$env:USERPROFILE\Downloads\ManageEngine_MDMLaptopEnrollment"
                if (-not (Test-Path $mdmDest)) {
                    $mdmSource = "D:\SOFTWARE\PAYOO\ManageEngine_MDMLaptopEnrollment"
                    if (Test-Path $mdmSource) {
                        Add-Status "Copying MDM files..."
                        try {
                            Copy-Item -Path $mdmSource -Destination $mdmDest -Recurse -Force -ErrorAction Stop
                            Add-Status "MDM           has been copied successfully!"
                        }
                        catch {
                            Add-Status "Error copying MDM: $_"
                        }
                    }
                    else {
                        Add-Status "Warning: MDM source folder not found at $mdmSource"
                    }
                }
                else {
                    Add-Status "MDM           is already copied. Skipping..."
                }
            }
            
            Add-Status "All files have been copied successfully."
            return $true
        }
        catch {
            Add-Status "CRITICAL ERROR in Copy-SoftwareFiles: $_"
            Add-Status "Error details: $($_.Exception.Message)"
            return $false
        }
    }

    # Install-Software function (DEFINED INSIDE THE DIALOG FUNCTION)
    function Install-Software {
        param ([string]$deviceType)

        try {
            $tempDir = "$env:USERPROFILE\Downloads\SETUP"
            $setupDir = "$tempDir\Software"
            $office2019Dir = "$tempDir\Office2019"
            
            # 1. Check and uninstall OneDrive if present
            $oneDrivePath = "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDriveUninstaller.exe"
            if (Test-Path $oneDrivePath) {
                Add-Status "OneDrive found. Uninstalling..."
                try {
                    Start-Process -FilePath $oneDrivePath -ArgumentList "/uninstall" -Wait -NoNewWindow
                    Add-Status "OneDrive uninstalled successfully!"
                }
                catch {
                    Add-Status "Warning: OneDrive uninstall failed: $_"
                }
            }
            else {
                Add-Status "OneDrive:     Has Not installed. Skipping..."
            }
            
            # 2. Install 7-Zip
            if (-not (Test-Path "C:\Program Files\7-Zip\7z.exe")) {
                $sevenZipInstaller = "$setupDir\7z2201-x64.msi"
                if (Test-Path $sevenZipInstaller) {
                    Add-Status "Installing 7-Zip..."
                    try {
                        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$sevenZipInstaller`" /quiet" -Wait
                        Add-Status "7-Zip installed successfully!"
                    }
                    catch {
                        Add-Status "ERROR: 7-Zip installation failed: $_"
                    }
                }
                else {
                    Add-Status "ERROR: 7-Zip installer not found at $sevenZipInstaller"
                }
            }
            else {
                Add-Status "7-Zip:        Already installed. Skipping..."
            }
            
            # 3. Install Chrome
            $chromeCheck = @(
                "C:\Program Files\Google\Chrome\Application\chrome.exe",
                "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            )
            $chromeInstalled = $false
            foreach ($path in $chromeCheck) {
                if (Test-Path $path) {
                    $chromeInstalled = $true
                    break
                }
            }
            
            if (-not $chromeInstalled) {
                $chromeInstaller = "$setupDir\ChromeSetup.exe"
                if (Test-Path $chromeInstaller) {
                    Add-Status "Installing Chrome..."
                    try {
                        Start-Process -FilePath $chromeInstaller -ArgumentList "/silent /install" -Wait
                        Add-Status "Chrome installed successfully!"
                    }
                    catch {
                        Add-Status "ERROR: Chrome installation failed: $_"
                    }
                }
                else {
                    Add-Status "ERROR: Chrome installer not found at $chromeInstaller"
                }
            }
            else {
                Add-Status "Chrome:       Already installed. Skipping..."
            }
            
            # 4. Install LAPS
            if (-not (Test-Path "C:\Program Files\LAPS\CSE\AdmPwd.dll")) {
                $lapsInstaller = "$setupDir\LAPS.x64.msi"
                if (Test-Path $lapsInstaller) {
                    Add-Status "Installing LAPS..."
                    try {
                        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$lapsInstaller`" /quiet" -Wait
                        Add-Status "LAPS installed successfully!"
                    }
                    catch {
                        Add-Status "ERROR: LAPS installation failed: $_"
                    }
                }
                else {
                    Add-Status "ERROR: LAPS installer not found at $lapsInstaller"
                }
            }
            else {
                Add-Status "LAPS:         Already installed. Skipping..."
            }
            
            # 5. Install Foxit Reader
            $foxitCheck = @(
                "C:\Program Files (x86)\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe",
                "C:\Program Files\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe"
            )
            $foxitInstalled = $false
            foreach ($path in $foxitCheck) {
                if (Test-Path $path) {
                    $foxitInstalled = $true
                    break
                }
            }
            
            if (-not $foxitInstalled) {
                $foxitInstaller = "$setupDir\FoxitPDFReader*.exe"
                $foxitFiles = Get-ChildItem -Path $setupDir -Name "FoxitPDFReader*.exe" -ErrorAction SilentlyContinue
                if ($foxitFiles.Count -gt 0) {
                    $foxitPath = "$setupDir\$($foxitFiles[0])"
                    Add-Status "Installing Foxit Reader..."
                    try {
                        Start-Process -FilePath $foxitPath -ArgumentList "/verysilent" -Wait
                        Add-Status "Foxit Reader installed successfully!"
                    }
                    catch {
                        Add-Status "ERROR: Foxit Reader installation failed: $_"
                    }
                }
                else {
                    Add-Status "ERROR: Foxit Reader installer not found in $setupDir"
                }
            }
            else {
                Add-Status "Foxit Reader: Already installed. Skipping..."
            }
            
            # 6. Install Office 2019
            if (-not (Test-Path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")) {
                $officeSetup = "$office2019Dir\setup.exe"
                if (Test-Path $officeSetup) {
                    Add-Status "Installing Office 2019..."
                    try {
                        Start-Process -FilePath $officeSetup -ArgumentList "/configure `"$office2019Dir\configuration.xml`"" -Wait
                        Add-Status "Office 2019 installed successfully!"
                    }
                    catch {
                        Add-Status "ERROR: Office 2019 installation failed: $_"
                    }
                }
                else {
                    Add-Status "ERROR: Office 2019 setup not found at $officeSetup"
                }
            }
            else {
                Add-Status "Office 2019:  Already installed. Skipping..."
            }
            
            # 7. Install Zoom
            if ($deviceType -eq "Laptop") {
                $zoomCheck = @(
                    "$env:USERPROFILE\AppData\Roaming\Zoom\bin\Zoom.exe",
                    "C:\Program Files\Zoom\bin\Zoom.exe",
                    "C:\Program Files (x86)\Zoom\bin\Zoom.exe"
                )
                $zoomInstalled = $false
                foreach ($path in $zoomCheck) {
                    if (Test-Path $path) {
                        $zoomInstalled = $true
                        break
                    }
                }
                if (-not $zoomInstalled) {
                    $zoomInstaller = "$setupDir\ZoomInstallerFull.exe"
                    if (Test-Path $zoomInstaller) {
                        Add-Status "Installing Zoom..."
                        try {
                            Start-Process -FilePath $zoomInstaller -ArgumentList "/silent" -Wait
                            Add-Status "Zoom installed successfully!"
                        }
                        catch {
                            Add-Status "ERROR: Zoom installation failed: $_"
                        }
                    }
                    else {
                        Add-Status "ERROR: Zoom installer not found at $zoomInstaller"
                    }
                }
                else {
                    Add-Status "Zoom:         Already installed. Skipping..."
                }
                
                # 8. Install CheckPointVPN
                if (-not (Test-Path "C:\Program Files (x86)\CheckPoint\Endpoint Connect\trac.exe")) {
                    $vpnInstaller = "$setupDir\CheckPointVPN.msi"
                    if (Test-Path $vpnInstaller) {
                        Add-Status "Installing CheckPointVPN..."
                        try {
                            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$vpnInstaller`" /quiet" -Wait
                            Add-Status "CheckPointVPN installed successfully!"
                        }
                        catch {
                            Add-Status "ERROR: CheckPointVPN installation failed: $_"
                        }
                    }
                    else {
                        Add-Status "ERROR: CheckPointVPN installer not found at $vpnInstaller"
                    }
                }
                else {
                    Add-Status "CheckPointVPN:Already installed. Skipping..."
                }
            }
            return $true
        }
        catch {
            Add-Status "CRITICAL ERROR in Install-Software: $_"
            Add-Status "Error details: $($_.Exception.Message)"
            return $false
        }
    }

    # Desktop button
    $btnDesktop = New-DynamicButton -text "DESKTOP" -x 10 -y 60 -width 200 -height 50 -clickAction {
        Add-Status "STEP 1: Copying required files for Desktop..."
        $copyResult = Copy-SoftwareFiles -deviceType "Desktop"

        if ($copyResult) {
            Add-Status "STEP 2: Installing software for Desktop..."
            $installResult = Install-Software -deviceType "Desktop"

            if ($installResult) {
                Add-Status "All software installation completed successfully!"
            }
            else {
                Add-Status "Warning: Some installations may have failed."
            }
        }
        else {
            Add-Status "Error: Failed to copy required files. Installation aborted."
        }
    }
    $deviceTypeForm.Controls.Add($btnDesktop)

    # Laptop button
    $btnLaptop = New-DynamicButton -text "LAPTOP" -x 260 -y 60 -width 200 -height 50 -clickAction {
        Add-Status "STEP 1: Copying required files for Laptop..."
        $copyResult = Copy-SoftwareFiles -deviceType "Laptop"

        if ($copyResult) {
            Add-Status "STEP 2: Installing software for Laptop..."
            $installResult = Install-Software -deviceType "Laptop"

            if ($installResult) {
                Add-Status "All software installation completed successfully!"
            }
            else {
                Add-Status "Warning: Some installations may have failed."
            }
        }
        else {
            Add-Status "Error: Failed to copy required files. Installation aborted."
        }
    }
    $deviceTypeForm.Controls.Add($btnLaptop)

    # Add KeyDown event handler for Esc key
    $deviceTypeForm.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $deviceTypeForm.Close()
        }
    })

    # Enable key events
    $deviceTypeForm.KeyPreview = $true

    # When form closes, show main menu
    $deviceTypeForm.Add_FormClosed({
        Show-MainMenu
    })

    # Show the dialog
    $deviceTypeForm.ShowDialog()
}

# [2] Install Software Button
$buttonInstallSoftware = New-DynamicButton -text "[2] Install All Software" -x 30 -y 180 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    Show-InstallSoftwareDialog
}

#================================================
# [3] Power Options
$buttonPowerOptions = New-DynamicButton -text "[3] Power Options" -x 30 -y 260 -width 380 -height 60 -clickAction {
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

            # Create a process to run the command with elevated privileges
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-Command Start-Process cmd.exe -ArgumentList '/c $command' -Verb RunAs -WindowStyle Hidden"
            $psi.UseShellExecute = $true
            $psi.Verb = "runas"
            $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

            # Start the process
            [System.Diagnostics.Process]::Start($psi)

            Add-Status "Time zone, power options have been configured successfully!"
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

# SECTION [4.3]: RENAME VOLUME FUNCTIONS (DEFINED BEFORE USE)
# Hàm tạo title cho form rename volume
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

# Hàm tạo groupbox cho form rename volume
function New-RenameVolumeGroupBox {
    param([System.Windows.Forms.Panel]$parentPanel, [System.Windows.Forms.ListBox]$driveListBox)
    $groupBox = New-Object System.Windows.Forms.GroupBox
    $groupBox.Text = "Volume Rename Configuration"
    $groupBox.Location = New-Object System.Drawing.Point(180, 60)
    $groupBox.Size = New-Object System.Drawing.Size(400, 150)
    $groupBox.ForeColor = [System.Drawing.Color]::Lime
    $groupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $parentPanel.Controls.Add($groupBox)

    # Drive letter label
    $driveLetterLabel = New-Object System.Windows.Forms.Label
    $driveLetterLabel.Text = "Drive Letter:"
    $driveLetterLabel.Location = New-Object System.Drawing.Point(30, 30)
    $driveLetterLabel.Size = New-Object System.Drawing.Size(100, 20)
    $driveLetterLabel.ForeColor = [System.Drawing.Color]::White
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
        GroupBox = $groupBox
        DriveLetterTextBox = $script:renameDriveLetterTextBox
        NewLabelTextBox = $script:renameNewLabelTextBox
    }
}

# Hàm tạo nút rename cho form rename volume
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
                } catch {
                    Add-Status "Error renaming drive: $_"
                }
            } else {
                Add-Status "Please enter both drive letter and new label."
            }
        }
    }
    $groupBox.Controls.Add($renameButton)
}

# SECTION [4.4]: EXTEND VOLUME FUNCTIONS - Các hàm mở rộng ổ đĩa
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

# [4] Change / Edit Volume
$buttonChangeVolume = New-DynamicButton -text "[4] Change / Edit Volume" -x 30 -y 340 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    # Hide the main menu
    Hide-MainMenu
    # Create volume management form
    $volumeForm = New-Object System.Windows.Forms.Form
    $volumeForm.Text = "Volume Management"
    $volumeForm.Size = New-Object System.Drawing.Size(820, 650) # Increase the size of the form
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
                if ($script:renameDriveLetterTextBox) {
                    $script:renameDriveLetterTextBox.Text = $driveLetter
                }
            }
        }
    })

    # [4.1] Change Drive Letter button
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

    # [4.2] Shrink Volume button
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

    # [4.3] Rename Volume button
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

    # [4.4] Extend Volume button
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

    # [4.0] Return to Main Menu button
    $btnReturn = New-DynamicButton -text "Return" -x 660 -y 150 -width 120 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        $volumeForm.Close()
    }
    $volumeForm.Controls.Add($btnReturn)

    # When the form is closed, show the main menu again
    $volumeForm.Add_FormClosed({
        Show-MainMenu
    })

    # Add KeyDown event handler for Esc key in volume form
    $volumeForm.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $volumeForm.Close()
        }
    })

    # Enable key events in volume form
    $volumeForm.KeyPreview = $true

    # Show the form
    $volumeForm.ShowDialog()
}

# [5] Activate Windows 10 Pro and Office 2019 Pro Plus
$buttonActivate = New-DynamicButton -text "[5] Activate" -x 30 -y 420 -width 380 -height 60 -clickAction { 
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
# [6] Turn On Features
$buttonTurnOnFeatures = New-DynamicButton -text "[6] Turn On Features" -x 430 -y 100 -width 380 -height 60 -clickAction { 
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

#================================================
# SECTION 7: RENAME DEVICE FUNCTIONS - Các hàm đổi tên máy tính
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

# [7] Rename Device
$buttonRenameDevice = New-DynamicButton -text "[7] Rename Device" -x 430 -y 180 -width 380 -height 60 -clickAction {
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

#================================================
# SECTION 8: SET PASSWORD FUNCTIONS - Các hàm đặt mật khẩu
# [8] Set Password
$buttonSetPassword = New-DynamicButton -text "[8] Set Password" -x 430 -y 260 -width 380 -height 60 -clickAction {
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

#================================================
# SECTION 9: DOMAIN MANAGEMENT FUNCTIONS - Các hàm quản lý domain
# Domain Management Configuration
$script:DomainConfig = @{
    FormWidth = 500
    FormHeight = 450
    FormHeightMinimal = 380
    ButtonY = 350
    ButtonYMinimal = 280
    ControlSpacing = 40
    DefaultWorkgroup = "WORKGROUP"
}

# Domain Management Helper Functions
function Get-ComputerDomainInfo {
    <#
    .SYNOPSIS
    Retrieves current computer's domain/workgroup information
    .DESCRIPTION
    Gets computer name, domain/workgroup name, and domain membership status
    .OUTPUTS
    Hashtable containing computer information
    #>
    try {
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop
        return @{
            ComputerName = $env:COMPUTERNAME
            Domain = $computerSystem.Domain
            IsPartOfDomain = $computerSystem.PartOfDomain
            Success = $true
        }
    }
    catch {
        Write-Warning "Failed to retrieve computer domain information: $_"
        return @{
            ComputerName = $env:COMPUTERNAME
            Domain = "Unknown"
            IsPartOfDomain = $false
            Success = $false
        }
    }
}

# Hàm tạo label cho form domain management
function New-DomainManagementLabel {
    <#
    .SYNOPSIS
    Creates a standardized label for the domain management form
    .PARAMETER Text
    The text to display on the label
    .PARAMETER X
    X coordinate position
    .PARAMETER Y
    Y coordinate position
    .PARAMETER Width
    Width of the label
    .PARAMETER Height
    Height of the label
    .PARAMETER FontSize
    Font size (default: 12)
    .PARAMETER FontStyle
    Font style (default: Regular)
    #>
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [int]$FontSize = 12,
        [System.Drawing.FontStyle]$FontStyle = [System.Drawing.FontStyle]::Regular
    )
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Font = New-Object System.Drawing.Font("Arial", $FontSize, $FontStyle)
    $label.ForeColor = [System.Drawing.Color]::White
    $label.Size = New-Object System.Drawing.Size($Width, $Height)
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    
    return $label
}

# Hàm tạo textbox cho form domain management
function New-DomainManagementTextBox {
    <#
    .SYNOPSIS
    Creates a standardized textbox for the domain management form
    .PARAMETER X
    X coordinate position
    .PARAMETER Y
    Y coordinate position
    .PARAMETER Width
    Width of the textbox
    .PARAMETER Height
    Height of the textbox
    .PARAMETER IsPassword
    Whether this is a password field
    .PARAMETER DefaultText
    Default text value
    #>
    param(
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [bool]$IsPassword = $false,
        [string]$DefaultText = ""
    )
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $textBox.Size = New-Object System.Drawing.Size($Width, $Height)
    $textBox.Location = New-Object System.Drawing.Point($X, $Y)
    $textBox.BackColor = [System.Drawing.Color]::White
    $textBox.ForeColor = [System.Drawing.Color]::Black
    $textBox.Text = $DefaultText
    
    if ($IsPassword) {
        $textBox.UseSystemPasswordChar = $true
    }
    
    return $textBox
}

# Hàm tạo radio button cho form domain management
function New-DomainManagementRadioButton {
    <#
    .SYNOPSIS
    Creates a standardized radio button for the domain management form
    .PARAMETER Text
    The text to display next to the radio button
    .PARAMETER X
    X coordinate position
    .PARAMETER Y
    Y coordinate position
    .PARAMETER Width
    Width of the radio button
    .PARAMETER Height
    Height of the radio button
    .PARAMETER IsChecked
    Whether the button is initially checked
    .PARAMETER IsEnabled
    Whether the button is enabled
    #>
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [bool]$IsChecked = $false,
        [bool]$IsEnabled = $true
    )
    
    $radioButton = New-Object System.Windows.Forms.RadioButton
    $radioButton.Text = $Text
    $radioButton.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioButton.ForeColor = [System.Drawing.Color]::White
    $radioButton.Location = New-Object System.Drawing.Point($X, $Y)
    $radioButton.Size = New-Object System.Drawing.Size($Width, $Height)
    $radioButton.BackColor = [System.Drawing.Color]::Black
    $radioButton.Checked = $IsChecked
    $radioButton.Enabled = $IsEnabled
    
    return $radioButton
}

# Hàm cập nhật layout cho form domain management
function Set-DomainFormLayout {
    <#
    .SYNOPSIS
    Updates the domain form layout based on the selected operation type
    .PARAMETER FormControls
    Hashtable containing all form controls
    .PARAMETER OperationType
    The type of operation: 'Domain', 'Workgroup', or 'LeaveDomain'
    #>
    param(
        [hashtable]$FormControls,
        [string]$OperationType
    )
    
    switch ($OperationType) {
        'Domain' {
            $FormControls.NameLabel.Text = "Domain Name:"
            $FormControls.NameTextBox.Text = ""
            $FormControls.UsernameLabel.Visible = $true
            $FormControls.UsernameTextBox.Visible = $true
            $FormControls.PasswordLabel.Visible = $true
            $FormControls.PasswordTextBox.Visible = $true
            $FormControls.JoinButton.Text = "Join"
            $FormControls.JoinButton.Location = New-Object System.Drawing.Point(30, $script:DomainConfig.ButtonY)
            $FormControls.CancelButton.Location = New-Object System.Drawing.Point(250, $script:DomainConfig.ButtonY)
            $FormControls.Form.Size = New-Object System.Drawing.Size($script:DomainConfig.FormWidth, $script:DomainConfig.FormHeight)
        }
        'Workgroup' {
            $FormControls.NameLabel.Text = "Workgroup Name:"
            $FormControls.NameTextBox.Text = $script:DomainConfig.DefaultWorkgroup
            $FormControls.UsernameLabel.Visible = $false
            $FormControls.UsernameTextBox.Visible = $false
            $FormControls.PasswordLabel.Visible = $false
            $FormControls.PasswordTextBox.Visible = $false
            $FormControls.JoinButton.Text = "Join"
            $FormControls.JoinButton.Location = New-Object System.Drawing.Point(30, $script:DomainConfig.ButtonYMinimal)
            $FormControls.CancelButton.Location = New-Object System.Drawing.Point(250, $script:DomainConfig.ButtonYMinimal)
            $FormControls.Form.Size = New-Object System.Drawing.Size($script:DomainConfig.FormWidth, $script:DomainConfig.FormHeightMinimal)
        }
        'LeaveDomain' {
            $FormControls.NameLabel.Text = "New Workgroup Name:"
            $FormControls.NameTextBox.Text = $script:DomainConfig.DefaultWorkgroup
            $FormControls.UsernameLabel.Visible = $false
            $FormControls.UsernameTextBox.Visible = $false
            $FormControls.PasswordLabel.Visible = $false
            $FormControls.PasswordTextBox.Visible = $false
            $FormControls.JoinButton.Text = "Leave Domain"
            $FormControls.JoinButton.Location = New-Object System.Drawing.Point(30, $script:DomainConfig.ButtonYMinimal)
            $FormControls.CancelButton.Location = New-Object System.Drawing.Point(250, $script:DomainConfig.ButtonYMinimal)
            $FormControls.Form.Size = New-Object System.Drawing.Size($script:DomainConfig.FormWidth, $script:DomainConfig.FormHeightMinimal)
        }
    }
}

# Hàm kiểm tra đầu vào cho join domain
function Test-DomainJoinInputs {
    <#
    .SYNOPSIS
    Validates inputs for domain join operation
    .PARAMETER DomainName
    The domain name to join
    .PARAMETER Username
    The username for domain authentication
    .PARAMETER Password
    The password for domain authentication
    .OUTPUTS
    Hashtable with validation result and error message
    #>
    param(
        [string]$DomainName,
        [string]$Username,
        [string]$Password
    )
    
    if ([string]::IsNullOrWhiteSpace($DomainName)) {
        return @{
            IsValid = $false
            ErrorMessage = "Domain name cannot be empty."
        }
    }
    
    if ([string]::IsNullOrWhiteSpace($Username)) {
        return @{
            IsValid = $false
            ErrorMessage = "Username is required for domain join."
        }
    }
    
    if ([string]::IsNullOrWhiteSpace($Password)) {
        return @{
            IsValid = $false
            ErrorMessage = "Password is required for domain join."
        }
    }
    
    # Additional domain name format validation
    if ($DomainName -notmatch '^[a-zA-Z0-9.-]+$') {
        return @{
            IsValid = $false
            ErrorMessage = "Domain name contains invalid characters. Use only letters, numbers, dots, and hyphens."
        }
    }
    
    return @{
        IsValid = $true
        ErrorMessage = ""
    }
}

# Hàm kiểm tra đầu vào cho join workgroup
function Test-WorkgroupInputs {
    <#
    .SYNOPSIS
    Validates inputs for workgroup join operation
    .PARAMETER WorkgroupName
    The workgroup name to join
    .OUTPUTS
    Hashtable with validation result and error message
    #>
    param(
        [string]$WorkgroupName
    )
    
    if ([string]::IsNullOrWhiteSpace($WorkgroupName)) {
        return @{
            IsValid = $false
            ErrorMessage = "Workgroup name cannot be empty."
        }
    }
    
    # Workgroup name validation (NetBIOS naming rules)
    if ($WorkgroupName.Length -gt 15) {
        return @{
            IsValid = $false
            ErrorMessage = "Workgroup name cannot exceed 15 characters."
        }
    }
    
    if ($WorkgroupName -match '[\\/:*?"<>|]') {
        return @{
            IsValid = $false
            ErrorMessage = "Workgroup name contains invalid characters."
        }
    }
    
    return @{
        IsValid = $true
        ErrorMessage = ""
    }
}

# Hàm thực hiện lệnh domain với quyền admin
function Invoke-ElevatedDomainCommand {
    <#
    .SYNOPSIS
    Executes a domain-related PowerShell command with elevated privileges
    .PARAMETER Command
    The PowerShell command to execute
    .PARAMETER OperationType
    The type of operation for user feedback
    .OUTPUTS
    Boolean indicating success
    #>
    param(
        [string]$Command,
        [string]$OperationType
    )
    
    try {
        # Create process start info for elevated execution
        $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processStartInfo.FileName = "powershell.exe"
        $processStartInfo.Arguments = "-Command Start-Process powershell.exe -ArgumentList '-Command $Command' -Verb RunAs"
        $processStartInfo.UseShellExecute = $true
        $processStartInfo.Verb = "runas"

        # Start the elevated process
        $process = [System.Diagnostics.Process]::Start($processStartInfo)
        
        if ($null -eq $process) {
            throw "Failed to start elevated process"
        }
        
        # Show appropriate success message
        $successMessages = @{
            'DomainJoin' = "Domain join command has been initiated. If prompted, please allow the elevation request. Your computer will restart to apply the changes."
            'WorkgroupJoin' = "Workgroup join command has been initiated. If prompted, please allow the elevation request. Your computer will restart to apply the changes."
            'LeaveDomain' = "Leave domain command has been initiated. If prompted, please allow the elevation request. Your computer will restart to apply the changes."
        }
        
        $message = $successMessages[$OperationType]
        if ([string]::IsNullOrEmpty($message)) {
            $message = "Command has been initiated. Your computer will restart to apply the changes."
        }
        
        [System.Windows.Forms.MessageBox]::Show(
            $message,
            $OperationType,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        return $true
    }
    catch {
        Write-Error "Failed to execute elevated domain command: $_"
        [System.Windows.Forms.MessageBox]::Show(
            "Error processing $OperationType operation: $_`n`nNote: This operation requires administrative privileges.",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}

# Hàm xử lý join domain
function Invoke-DomainJoinOperation {
    param(
        [string]$DomainName,
        [string]$Username,
        [string]$Password
    )
    
    # Validate inputs
    $validation = Test-DomainJoinInputs -DomainName $DomainName -Username $Username -Password $Password
    if (-not $validation.IsValid) {
        [System.Windows.Forms.MessageBox]::Show(
            $validation.ErrorMessage,
            "Validation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
    
    # Escape special characters in password for command line
    $escapedPassword = $Password -replace "'", "''"
    
    # Build domain join command
    $command = "Add-Computer -DomainName '$DomainName' -Credential (New-Object System.Management.Automation.PSCredential ('$Username', (ConvertTo-SecureString '$escapedPassword' -AsPlainText -Force))) -Restart -Force"
    
    return Invoke-ElevatedDomainCommand -Command $command -OperationType "DomainJoin"
}

# Hàm xử lý join workgroup
function Invoke-WorkgroupJoinOperation {
    <#
    .SYNOPSIS
    Handles workgroup join operation
    .PARAMETER WorkgroupName
    The workgroup name to join
    .OUTPUTS
    Boolean indicating success
    #>
    param(
        [string]$WorkgroupName
    )
    
    # Validate inputs
    $validation = Test-WorkgroupInputs -WorkgroupName $WorkgroupName
    if (-not $validation.IsValid) {
        [System.Windows.Forms.MessageBox]::Show(
            $validation.ErrorMessage,
            "Validation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
    
    # Build workgroup join command
    $command = "Add-Computer -WorkgroupName '$WorkgroupName' -Restart -Force"
    
    return Invoke-ElevatedDomainCommand -Command $command -OperationType "WorkgroupJoin"
}

# Hàm xử lý rời domain
function Invoke-LeaveDomainOperation {
    <#
    .SYNOPSIS
    Handles leave domain operation
    .PARAMETER NewWorkgroupName
    The new workgroup name to join after leaving domain
    .OUTPUTS
    Boolean indicating success
    #>
    param(
        [string]$NewWorkgroupName
    )
    
    # Validate inputs
    $validation = Test-WorkgroupInputs -WorkgroupName $NewWorkgroupName
    if (-not $validation.IsValid) {
        [System.Windows.Forms.MessageBox]::Show(
            $validation.ErrorMessage,
            "Validation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
    
    # Confirm leave domain operation
    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to leave the current domain and join the workgroup '$NewWorkgroupName'? Your computer will restart after this operation.",
        "Confirm Leave Domain",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return $false
    }
    
    # Build leave domain command
    $command = "Remove-Computer -WorkgroupName '$NewWorkgroupName' -Force -Restart"
    
    return Invoke-ElevatedDomainCommand -Command $command -OperationType "LeaveDomain"
}

# Hàm hiển thị form domain management
function Show-DomainManagementForm {
    <#
    .SYNOPSIS
    Creates and displays the domain management form
    .DESCRIPTION
    Main function that orchestrates the domain management UI
    #>
    
    # Hide the main menu
    Hide-MainMenu
    
    # Get current computer information
    $computerInfo = Get-ComputerDomainInfo
    
    # Create main form
    $joinForm = New-Object System.Windows.Forms.Form
    $joinForm.Text = "Domain Management"
    $joinForm.Size = New-Object System.Drawing.Size($script:DomainConfig.FormWidth, $script:DomainConfig.FormHeight)
    $joinForm.StartPosition = "CenterScreen"
    $joinForm.BackColor = [System.Drawing.Color]::Black
    $joinForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $joinForm.MaximizeBox = $false
    $joinForm.MinimizeBox = $false

    # Create title label
    $titleLabel = New-DomainManagementLabel -Text "DOMAIN MANAGEMENT" -X 10 -Y 20 -Width 480 -Height 40 -FontSize 14 -FontStyle ([System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $joinForm.Controls.Add($titleLabel)

    # Current computer info labels
    $currentLabel = New-DomainManagementLabel -Text "Current Computer Name: $($computerInfo.ComputerName)" -X 20 -Y 70 -Width 480 -Height 30
    $joinForm.Controls.Add($currentLabel)

    $domainStatusText = if ($computerInfo.IsPartOfDomain) {
        "Currently joined to DOMAIN: $($computerInfo.Domain)"
    } else {
        "Currently joined to WORKGROUP: $($computerInfo.Domain)"
    }
    $domainLabel = New-DomainManagementLabel -Text $domainStatusText -X 20 -Y 100 -Width 480 -Height 30
    $joinForm.Controls.Add($domainLabel)

    # Create radio buttons group
    $groupBox = New-Object System.Windows.Forms.GroupBox
    $groupBox.Text = "Select Option"
    $groupBox.Font = New-Object System.Drawing.Font("Arial", 10)
    $groupBox.ForeColor = [System.Drawing.Color]::White
    $groupBox.Size = New-Object System.Drawing.Size(460, 80)
    $groupBox.Location = New-Object System.Drawing.Point(20, 140)
    $groupBox.BackColor = [System.Drawing.Color]::Black

    $radioDomain = New-DomainManagementRadioButton -Text "Join Domain" -X 20 -Y 30 -Width 120 -Height 30 -IsChecked $true
    $radioWorkgroup = New-DomainManagementRadioButton -Text "Join Workgroup" -X 150 -Y 30 -Width 140 -Height 30
    $radioLeaveDomain = New-DomainManagementRadioButton -Text "Leave Domain" -X 300 -Y 30 -Width 140 -Height 30 -IsEnabled $computerInfo.IsPartOfDomain

    $groupBox.Controls.Add($radioDomain)
    $groupBox.Controls.Add($radioWorkgroup)
    $groupBox.Controls.Add($radioLeaveDomain)
    $joinForm.Controls.Add($groupBox)

    # Create input controls
    $nameLabel = New-DomainManagementLabel -Text "Domain Name:" -X 20 -Y 230 -Width 150 -Height 30
    $nameTextBox = New-DomainManagementTextBox -X 170 -Y 230 -Width 300 -Height 30
    $usernameLabel = New-DomainManagementLabel -Text "Username:" -X 20 -Y 270 -Width 150 -Height 30
    $usernameTextBox = New-DomainManagementTextBox -X 170 -Y 270 -Width 300 -Height 30
    $passwordLabel = New-DomainManagementLabel -Text "Password:" -X 20 -Y 310 -Width 150 -Height 30
    $passwordTextBox = New-DomainManagementTextBox -X 170 -Y 310 -Width 300 -Height 30 -IsPassword $true

    $joinForm.Controls.AddRange(@($nameLabel, $nameTextBox, $usernameLabel, $usernameTextBox, $passwordLabel, $passwordTextBox))

    # Create buttons
    $joinButton = New-DynamicButton -text "Join" -x 30 -y $script:DomainConfig.ButtonY -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 180, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 220, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 140, 0)) -textColor ([System.Drawing.Color]::White) -fontSize 12 -fontStyle ([System.Drawing.FontStyle]::Bold)
    
    $cancelButton = New-DynamicButton -text "Cancel" -x 250 -y $script:DomainConfig.ButtonY -width 200 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
        $joinForm.Close()
    }

    $joinForm.Controls.AddRange(@($joinButton, $cancelButton))

    # Store form controls for easy access
    $formControls = @{
        Form = $joinForm
        NameLabel = $nameLabel
        NameTextBox = $nameTextBox
        UsernameLabel = $usernameLabel
        UsernameTextBox = $usernameTextBox
        PasswordLabel = $passwordLabel
        PasswordTextBox = $passwordTextBox
        JoinButton = $joinButton
        CancelButton = $cancelButton
    }

    # Event handlers for radio buttons
    $radioDomain.Add_CheckedChanged({
        if ($radioDomain.Checked) {
            Set-DomainFormLayout -FormControls $formControls -OperationType 'Domain'
        }
    })

    $radioWorkgroup.Add_CheckedChanged({
        if ($radioWorkgroup.Checked) {
            Set-DomainFormLayout -FormControls $formControls -OperationType 'Workgroup'
        }
    })

    $radioLeaveDomain.Add_CheckedChanged({
        if ($radioLeaveDomain.Checked) {
            Set-DomainFormLayout -FormControls $formControls -OperationType 'LeaveDomain'
        }
    })

    # Join button click handler
    $joinButton.Add_Click({
        $name = $nameTextBox.Text.Trim()
        $success = $false

        try {
            if ($radioDomain.Checked) {
                $success = Invoke-DomainJoinOperation -DomainName $name -Username $usernameTextBox.Text.Trim() -Password $passwordTextBox.Text
            }
            elseif ($radioWorkgroup.Checked) {
                $success = Invoke-WorkgroupJoinOperation -WorkgroupName $name
            }
            else {
                $success = Invoke-LeaveDomainOperation -NewWorkgroupName $name
            }

            if ($success) {
                $joinForm.Close()
            }
        }
        catch {
            Write-Error "Unexpected error in domain management operation: $_"
            [System.Windows.Forms.MessageBox]::Show(
                "An unexpected error occurred: $_",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    # Set form behavior
    $joinForm.AcceptButton = $joinButton
    $joinForm.CancelButton = $cancelButton
    $joinForm.Add_FormClosed({
        Show-MainMenu
    })

    # Show the form
    $joinForm.ShowDialog()
}
# [9] Join Domain
$buttonJoinDomain = New-DynamicButton -text "[9] Join Domain" -x 430 -y 340 -width 380 -height 60 -clickAction {
    Show-DomainManagementForm
}

#================================================
# [0] Exit
$buttonExit = New-DynamicButton -text "[0] Exit" -x 430 -y 420 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
    $script:form.Close()
}

# Add KeyDown event handler for Esc key
$script:form.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
        $script:form.Close()
    }
})

# Enable key events
$script:form.KeyPreview = $true

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