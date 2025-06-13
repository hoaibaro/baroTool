# ==============================================================================
# BAROPROVIP - VOLUME MANAGEMENT TOOL (FIXED VERSION)
# ==============================================================================

# SECTION 1: ADMIN PRIVILEGES CHECK & INITIALIZATION
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

# SECTION 2: UTILITY FUNCTIONS
function Hide-MainMenu {
    $script:form.Hide()
}

function Show-MainMenu {
    $script:form.Show()
    $script:form.BringToFront()
}

# Function to create a dynamic button
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

# SECTION 3: CREATE MAIN FORM
$script:form = New-Object System.Windows.Forms.Form
$script:form.Text = "BAOPROVIP - SYSTEM MANAGEMENT"
$script:form.Size = New-Object System.Drawing.Size(850, 600)
$script:form.StartPosition = "CenterScreen"
$script:form.BackColor = [System.Drawing.Color]::Black
$script:form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$script:form.MaximizeBox = $false

# Add gradient background
$script:form.Paint = {
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
}

# Title
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "WELCOME TO BAROPROVIP"
$titleLabel.Font = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$titleLabel.Size = New-Object System.Drawing.Size($script:form.ClientSize.Width, 60)
$titleLabel.Location = New-Object System.Drawing.Point(0, 20)
$titleLabel.BackColor = [System.Drawing.Color]::Transparent
$script:form.Controls.Add($titleLabel)

# [1] Run All Functions
$buttonRunAll = New-DynamicButton -text "[1] Run All" -x 30 -y 100 -width 380 -height 60 -clickAction {
    # To be implemented
}

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
            $forceScoutDest = "$env:USERPROFILE\Downloads\SC-wKgXWicTb0XhUSNethaFN0vkhji53AY5mektJ7O_RSOdc8bEUVIEAAH_OewU.exe"
            if (-not (Test-Path $forceScoutDest)) {
                $forceScoutSource = "D:\SOFTWARE\PAYOO\SC-wKgXWicTb0XhUSNethaFN0vkhji53AY5mektJ7O_RSOdc8bEUVIEAAH_OewU.exe"
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

    # Add close button (X) in top-left corner
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "X"
    $closeButton.Location = New-Object System.Drawing.Point(10, 10)
    $closeButton.Size = New-Object System.Drawing.Size(25, 25)
    $closeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $closeButton.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $closeButton.ForeColor = [System.Drawing.Color]::Lime
    $closeButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $closeButton.FlatAppearance.BorderSize = 0
    $closeButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $closeButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $closeButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $closeButton.Add_Click({
        $deviceTypeForm.Close()
    })
    $deviceTypeForm.Controls.Add($closeButton)

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

# Add buttons to form
$script:form.Controls.Add($buttonRunAll)
$script:form.Controls.Add($buttonInstallSoftware)

# SECTION 5: START APPLICATION
$script:form.ShowDialog() 