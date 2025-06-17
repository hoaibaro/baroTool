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

# Function to add status message
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

# Copy-SoftwareFiles function
function Copy-SoftwareFiles {
    param ([string]$deviceType, [System.Windows.Forms.TextBox]$statusTextBox)

    try {       
        $tempDir = "$env:USERPROFILE\Downloads\SETUP"
         
        if (-not (Test-Path $tempDir)) {
            Add-Status "Creating temporary folder..." $statusTextBox
            New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
            Add-Status "Temporary folder created successfully!" $statusTextBox
        }
        else {
            Add-Status "Temporary folder already exists. Skipping..." $statusTextBox
        }

        # Check D: drive
        if (-not (Test-Path "D:\")) {
            Add-Status "WARNING: D drive not found. Creating mock installation..." $statusTextBox
            
            if (-not (Test-Path "$tempDir\Software")) {
                New-Item -Path "$tempDir\Software" -ItemType Directory -Force | Out-Null
                Add-Status "Created mock Software directory" $statusTextBox
            }
             
            if (-not (Test-Path "$tempDir\Office2019")) {
                New-Item -Path "$tempDir\Office2019" -ItemType Directory -Force | Out-Null
                Add-Status "Created mock Office2019 directory" $statusTextBox
            }
             
            Add-Status "Copy-SoftwareFiles completed (mock mode)" $statusTextBox
            return $true
        }
         
        # Copy SETUP folder from D:\SOFTWARE\PAYOO\SETUP
        if (-not (Test-Path "$tempDir\Software")) {
            $setupSource = "D:\SOFTWARE\PAYOO\SETUP"
            if (Test-Path $setupSource) {
                Add-Status "Copying setup files from $setupSource..." $statusTextBox
                try {
                    Copy-Item -Path $setupSource -Destination "$tempDir\Software" -Recurse -Force -ErrorAction Stop
                    Add-Status "SetupFiles    has been copied successfully!" $statusTextBox
                }
                catch {
                    Add-Status "Error copying setup files: $_" $statusTextBox
                }
            }
            else {
                Add-Status "Warning: Setup source folder not found at $setupSource" $statusTextBox
            }
        }
        else {
            Add-Status "SetupFiles    is already copied. Skipping..." $statusTextBox
        }

        # Copy Office 2019
        if (-not (Test-Path "$tempDir\Office2019")) {
            $officeSource = "D:\SOFTWARE\OFFICE\Office 2019"
            if (Test-Path $officeSource) {
                Add-Status "Copying Office 2019 files from $officeSource..." $statusTextBox
                try {
                    New-Item -Path "$tempDir\Office2019" -ItemType Directory -Force | Out-Null
                    Copy-Item -Path "$officeSource\*" -Destination "$tempDir\Office2019" -Recurse -Force -ErrorAction Stop
                    Add-Status "Office 2019   has been copied successfully!" $statusTextBox
                }
                catch {
                    Add-Status "Error copying Office 2019: $_" $statusTextBox
                }
            }
            else {
                Add-Status "Warning: Office source folder not found at $officeSource" $statusTextBox
            }
        }
        else {
            Add-Status "Office 2019   is already copied. Skipping..." $statusTextBox
        }

        # Copy Unikey to C:\ drive
        if (-not (Test-Path "C:\unikey46RC2-230919-win64")) {
            $unikeySource = "D:\SOFTWARE\PAYOO\unikey46RC2-230919-win64"
            if (Test-Path $unikeySource) {
                Add-Status "Copying Unikey files to C:\ drive..." $statusTextBox
                try {
                    Copy-Item -Path $unikeySource -Destination "C:\unikey46RC2-230919-win64" -Recurse -Force -ErrorAction Stop
                    Add-Status "Unikey        has been copied successfully!" $statusTextBox
                }
                catch {
                    Add-Status "Error copying Unikey: $_" $statusTextBox
                }
            }
            else {
                Add-Status "Warning: Unikey source folder not found at $unikeySource" $statusTextBox
            }
        }
        else {
            Add-Status "Unikey        is already copied. Skipping..." $statusTextBox
        }

        # Copy MSTeamsSetup to C:\ drive
        if (-not (Test-Path "C:\MSTeamsSetup.exe")) {
            $teamsSource = "D:\SOFTWARE\PAYOO\MSTeamsSetup.exe"
            if (Test-Path $teamsSource) {
                Add-Status "Copying MSTeamsSetup file to C:\ drive..." $statusTextBox
                try {
                    Copy-Item -Path $teamsSource -Destination "C:\MSTeamsSetup.exe" -Force -ErrorAction Stop
                    Add-Status "MSTeamsSetup  has been copied successfully!" $statusTextBox
                }
                catch {
                    Add-Status "Error copying MSTeamsSetup: $_" $statusTextBox
                }
            }
            else {
                Add-Status "Warning: MSTeamsSetup source file not found at $teamsSource" $statusTextBox
            }
        }
        else {
            Add-Status "MSTeamsSetup  is already copied. Skipping..." $statusTextBox
        }

        # Copy ForceScout
        $forceScoutDest = "$env:USERPROFILE\Downloads\ForceScout.exe"
        if (-not (Test-Path $forceScoutDest)) {
            $forceScoutSource = "D:\SOFTWARE\PAYOO\ForceScout.exe"
            if (Test-Path $forceScoutSource) {
                Add-Status "Copying ForceScout file..." $statusTextBox
                try {
                    Copy-Item -Path $forceScoutSource -Destination $forceScoutDest -Force -ErrorAction Stop
                    Add-Status "ForceScout    has been copied successfully!" $statusTextBox
                }
                catch {
                    Add-Status "Error copying ForceScout: $_" $statusTextBox
                }
            }
            else {
                Add-Status "Warning: ForceScout source file not found at $forceScoutSource" $statusTextBox
            }
        }
        else {
            Add-Status "ForceScout    is already copied. Skipping..." $statusTextBox
        }

        # Copy FalconSensor folder
        $falconDest = "$env:USERPROFILE\Downloads\FalconSensor_Windows_installer (All AV)"
        if (-not (Test-Path $falconDest)) {
            $falconSource = "D:\SOFTWARE\PAYOO\FalconSensor_Windows_installer (All AV)"
            if (Test-Path $falconSource) {
                Add-Status "Copying FalconSensor folder..." $statusTextBox
                try {
                    Copy-Item -Path $falconSource -Destination $falconDest -Recurse -Force -ErrorAction Stop
                    Add-Status "FalconSensor  has been copied successfully!" $statusTextBox
                }
                catch {
                    Add-Status "Error copying FalconSensor: $_" $statusTextBox
                }
            }
            else {
                Add-Status "Warning: FalconSensor source folder not found at $falconSource" $statusTextBox
            }
        }
        else {
            Add-Status "FalconSensor  is already copied. Skipping..." $statusTextBox
        }

        # Copy device-specific agent
        if ($deviceType -eq "Desktop") {
            $agentDest = "$env:USERPROFILE\Downloads\Desktop Agent.exe"
            if (-not (Test-Path $agentDest)) {
                $agentSource = "D:\SOFTWARE\PAYOO\Desktop Agent.exe"
                if (Test-Path $agentSource) {
                    Add-Status "Copying Desktop Agent file..." $statusTextBox
                    try {
                        Copy-Item -Path $agentSource -Destination $agentDest -Force -ErrorAction Stop
                        Add-Status "Desktop Agent has been copied successfully!" $statusTextBox
                    }
                    catch {
                        Add-Status "Error copying Desktop Agent: $_" $statusTextBox
                    }
                }
                else {
                    Add-Status "Warning: Desktop Agent source file not found at $agentSource" $statusTextBox
                }
            }
            else {
                Add-Status "Desktop Agent is already copied. Skipping..." $statusTextBox
            }
        }
        elseif ($deviceType -eq "Laptop") {
            # Copy Laptop Agent
            $agentDest = "$env:USERPROFILE\Downloads\Laptop Agent.exe"
            if (-not (Test-Path $agentDest)) {
                $agentSource = "D:\SOFTWARE\PAYOO\Laptop Agent.exe"
                if (Test-Path $agentSource) {
                    Add-Status "Copying Laptop Agent file..." $statusTextBox
                    try {
                        Copy-Item -Path $agentSource -Destination $agentDest -Force -ErrorAction Stop
                        Add-Status "Laptop Agent  has been copied successfully!" $statusTextBox
                    }
                    catch {
                        Add-Status "Error copying Laptop Agent: $_" $statusTextBox
                    }
                }
                else {
                    Add-Status "Warning: Laptop Agent source file not found at $agentSource" $statusTextBox
                }
            }
            else {
                Add-Status "Laptop Agent  is already copied. Skipping..." $statusTextBox
            }

            # Copy MDM for laptops
            $mdmDest = "$env:USERPROFILE\Downloads\ManageEngine_MDMLaptopEnrollment"
            if (-not (Test-Path $mdmDest)) {
                $mdmSource = "D:\SOFTWARE\PAYOO\ManageEngine_MDMLaptopEnrollment"
                if (Test-Path $mdmSource) {
                    Add-Status "Copying MDM files..." $statusTextBox
                    try {
                        Copy-Item -Path $mdmSource -Destination $mdmDest -Recurse -Force -ErrorAction Stop
                        Add-Status "MDM           has been copied successfully!" $statusTextBox
                    }
                    catch {
                        Add-Status "Error copying MDM: $_" $statusTextBox
                    }
                }
                else {
                    Add-Status "Warning: MDM source folder not found at $mdmSource" $statusTextBox
                }
            }
            else {
                Add-Status "MDM           is already copied. Skipping..." $statusTextBox
            }
        }
        
        Add-Status "All files have been copied successfully." $statusTextBox
        return $true
    }
    catch {
        Add-Status "CRITICAL ERROR in Copy-SoftwareFiles: $_" $statusTextBox
        Add-Status "Error details: $($_.Exception.Message)" $statusTextBox
        return $false
    }
}

# Install-Software function
function Install-Software {
    param ([string]$deviceType, [System.Windows.Forms.TextBox]$statusTextBox)

    try {
        $tempDir = "$env:USERPROFILE\Downloads\SETUP"
        $setupDir = "$tempDir\Software"
        $office2019Dir = "$tempDir\Office2019"
        
        # 1. Check and uninstall OneDrive if present
        $oneDrivePath = "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDriveUninstaller.exe"
        if (Test-Path $oneDrivePath) {
            Add-Status "OneDrive found. Uninstalling..." $statusTextBox
            try {
                Start-Process -FilePath $oneDrivePath -ArgumentList "/uninstall" -Wait -NoNewWindow
                Add-Status "OneDrive uninstalled successfully!" $statusTextBox
            }
            catch {
                Add-Status "Warning: OneDrive uninstall failed: $_" $statusTextBox
            }
        }
        else {
            Add-Status "OneDrive:     Has Not installed. Skipping..." $statusTextBox
        }
        
        # 2. Install 7-Zip
        if (-not (Test-Path "C:\Program Files\7-Zip\7z.exe")) {
            $sevenZipInstaller = "$setupDir\7z2201-x64.msi"
            if (Test-Path $sevenZipInstaller) {
                Add-Status "Installing 7-Zip..." $statusTextBox
                try {
                    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$sevenZipInstaller`" /quiet" -Wait
                    Add-Status "7-Zip installed successfully!" $statusTextBox
                }
                catch {
                    Add-Status "ERROR: 7-Zip installation failed: $_" $statusTextBox
                }
            }
            else {
                Add-Status "ERROR: 7-Zip installer not found at $sevenZipInstaller" $statusTextBox
            }
        }
        else {
            Add-Status "7-Zip:        Already installed. Skipping..." $statusTextBox
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
                Add-Status "Installing Chrome..." $statusTextBox
                try {
                    Start-Process -FilePath $chromeInstaller -ArgumentList "/silent /install" -Wait
                    Add-Status "Chrome installed successfully!" $statusTextBox
                }
                catch {
                    Add-Status "ERROR: Chrome installation failed: $_" $statusTextBox
                }
            }
            else {
                Add-Status "ERROR: Chrome installer not found at $chromeInstaller" $statusTextBox
            }
        }
        else {
            Add-Status "Chrome:       Already installed. Skipping..." $statusTextBox
        }
        
        # 4. Install LAPS
        if (-not (Test-Path "C:\Program Files\LAPS\CSE\AdmPwd.dll")) {
            $lapsInstaller = "$setupDir\LAPS.x64.msi"
            if (Test-Path $lapsInstaller) {
                Add-Status "Installing LAPS..." $statusTextBox
                try {
                    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$lapsInstaller`" /quiet" -Wait
                    Add-Status "LAPS installed successfully!" $statusTextBox
                }
                catch {
                    Add-Status "ERROR: LAPS installation failed: $_" $statusTextBox
                }
            }
            else {
                Add-Status "ERROR: LAPS installer not found at $lapsInstaller" $statusTextBox
            }
        }
        else {
            Add-Status "LAPS:         Already installed. Skipping..." $statusTextBox
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
                Add-Status "Installing Foxit Reader..." $statusTextBox
                try {
                    Start-Process -FilePath $foxitPath -ArgumentList "/verysilent" -Wait
                    Add-Status "Foxit Reader installed successfully!" $statusTextBox
                }
                catch {
                    Add-Status "ERROR: Foxit Reader installation failed: $_" $statusTextBox
                }
            }
            else {
                Add-Status "ERROR: Foxit Reader installer not found in $setupDir" $statusTextBox
            }
        }
        else {
            Add-Status "Foxit Reader: Already installed. Skipping..." $statusTextBox
        }
        
        # 6. Install Office 2019
        if (-not (Test-Path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")) {
            $officeSetup = "$office2019Dir\setup.exe"
            if (Test-Path $officeSetup) {
                Add-Status "Installing Office 2019..." $statusTextBox
                try {
                    Start-Process -FilePath $officeSetup -ArgumentList "/configure `"$office2019Dir\configuration.xml`"" -Wait
                    Add-Status "Office 2019 installed successfully!" $statusTextBox
                }
                catch {
                    Add-Status "ERROR: Office 2019 installation failed: $_" $statusTextBox
                }
            }
            else {
                Add-Status "ERROR: Office 2019 setup not found at $officeSetup" $statusTextBox
            }
        }
        else {
            Add-Status "Office 2019:  Already installed. Skipping..." $statusTextBox
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
                    Add-Status "Installing Zoom..." $statusTextBox
                    try {
                        Start-Process -FilePath $zoomInstaller -ArgumentList "/silent" -Wait
                        Add-Status "Zoom installed successfully!" $statusTextBox
                    }
                    catch {
                        Add-Status "ERROR: Zoom installation failed: $_" $statusTextBox
                    }
                }
                else {
                    Add-Status "ERROR: Zoom installer not found at $zoomInstaller" $statusTextBox
                }
            }
            else {
                Add-Status "Zoom:         Already installed. Skipping..." $statusTextBox
            }
            
            # 8. Install CheckPointVPN
            if (-not (Test-Path "C:\Program Files (x86)\CheckPoint\Endpoint Connect\trac.exe")) {
                $vpnInstaller = "$setupDir\CheckPointVPN.msi"
                if (Test-Path $vpnInstaller) {
                    Add-Status "Installing CheckPointVPN..." $statusTextBox
                    try {
                        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$vpnInstaller`" /quiet" -Wait
                        Add-Status "CheckPointVPN installed successfully!" $statusTextBox
                    }
                    catch {
                        Add-Status "ERROR: CheckPointVPN installation failed: $_" $statusTextBox
                    }
                }
                else {
                    Add-Status "ERROR: CheckPointVPN installer not found at $vpnInstaller" $statusTextBox
                }
            }
            else {
                Add-Status "CheckPointVPN:Already installed. Skipping..." $statusTextBox
            }
        }
        return $true
    }
    catch {
        Add-Status "CRITICAL ERROR in Install-Software: $_" $statusTextBox
        Add-Status "Error details: $($_.Exception.Message)" $statusTextBox
        return $false
    }
}

# 
function Invoke-SystemConfiguration {
    param (
        [string]$deviceType,
        [System.Windows.Forms.TextBox]$statusTextBox
    )
    try {
        # --- Hiển thị tên máy tính hiện tại và đổi tên ---
        $currentName = $env:COMPUTERNAME
        Add-Status "Current computer name: $currentName" $statusTextBox
        
        # Tạo form hiển thị thông tin và nhập tên mới
        $renameForm = New-Object System.Windows.Forms.Form
        $renameForm.Text = "Computer Name Configuration"
        $renameForm.Size = New-Object System.Drawing.Size(450, 250)
        $renameForm.StartPosition = "CenterScreen"
        $renameForm.BackColor = [System.Drawing.Color]::Black
        $renameForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $renameForm.MaximizeBox = $false
        $renameForm.MinimizeBox = $false

        # Label hiển thị tên hiện tại
        $currentNameLabel = New-Object System.Windows.Forms.Label
        $currentNameLabel.Text = "Current Computer Name: $currentName"
        $currentNameLabel.Location = New-Object System.Drawing.Point(20, 20)
        $currentNameLabel.Size = New-Object System.Drawing.Size(400, 25)
        $currentNameLabel.ForeColor = [System.Drawing.Color]::White
        $currentNameLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $currentNameLabel.BackColor = [System.Drawing.Color]::Transparent
        $renameForm.Controls.Add($currentNameLabel)

        # Xác định prefix dựa trên loại thiết bị
        $prefix = ""
        if ($deviceType -eq "Desktop") {
            $prefix = "HOD"
        } elseif ($deviceType -eq "Laptop") {
            $prefix = "HOL"
        }

        # Label hướng dẫn nhập tên mới
        $instructionLabel = New-Object System.Windows.Forms.Label
        $instructionLabel.Text = "Enter new name (will be prefixed with $prefix):"
        $instructionLabel.Location = New-Object System.Drawing.Point(20, 60)
        $instructionLabel.Size = New-Object System.Drawing.Size(400, 25)
        $instructionLabel.ForeColor = [System.Drawing.Color]::Lime
        $instructionLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $instructionLabel.BackColor = [System.Drawing.Color]::Transparent
        $renameForm.Controls.Add($instructionLabel)

        # TextBox nhập tên mới
        $nameTextBox = New-Object System.Windows.Forms.TextBox
        $nameTextBox.Location = New-Object System.Drawing.Point(20, 90)
        $nameTextBox.Size = New-Object System.Drawing.Size(300, 25)
        $nameTextBox.Font = New-Object System.Drawing.Font("Arial", 10)
        $renameForm.Controls.Add($nameTextBox)

        # Label hiển thị preview tên mới
        $previewLabel = New-Object System.Windows.Forms.Label
        $previewLabel.Text = "New name will be: $prefix"
        $previewLabel.Location = New-Object System.Drawing.Point(20, 125)
        $previewLabel.Size = New-Object System.Drawing.Size(400, 25)
        $previewLabel.ForeColor = [System.Drawing.Color]::Yellow
        $previewLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Italic)
        $previewLabel.BackColor = [System.Drawing.Color]::Transparent
        $renameForm.Controls.Add($previewLabel)

        # Cập nhật preview khi người dùng gõ
        $nameTextBox.Add_TextChanged({
            $newPreview = $prefix + $nameTextBox.Text.Trim()
            $previewLabel.Text = "New name will be: $newPreview"
        })

        # Nút OK
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(250, 160)
        $okButton.Size = New-Object System.Drawing.Size(80, 30)
        $okButton.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 0)
        $okButton.ForeColor = [System.Drawing.Color]::White
        $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $okButton.Add_Click({
            $renameForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $renameForm.Close()
        })
        $renameForm.Controls.Add($okButton)

        # Nút Cancel
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(340, 160)
        $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
        $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
        $cancelButton.ForeColor = [System.Drawing.Color]::White
        $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $cancelButton.Add_Click({
            $renameForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $renameForm.Close()
        })
        $renameForm.Controls.Add($cancelButton)

        # Hiển thị form và xử lý kết quả
        $result = $renameForm.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $inputName = $nameTextBox.Text.Trim()
            
            if ($inputName -and $inputName -ne "") {
                $newName = $prefix + $inputName
                
                if ($newName -ne $currentName) {
                    Add-Status "Renaming computer from '$currentName' to '$newName'..." $statusTextBox
                    try {
                        Rename-Computer -NewName $newName -Force -ErrorAction Stop
                        Add-Status "Computer will be renamed to '$newName' after restart." $statusTextBox
                    } catch {
                        Add-Status "ERROR: Failed to rename computer: $_" $statusTextBox
                    }
                } else {
                    Add-Status "New name is same as current name. Skipping..." $statusTextBox
                }
            } else {
                Add-Status "No computer name entered. Skipping rename..." $statusTextBox
            }
        } else {
            Add-Status "Computer rename cancelled by user." $statusTextBox
        }

        # --- Tạo lối tắt trên Desktop ---
        $publicDesktop = "$env:PUBLIC\Desktop"
        Add-Status "Creating shortcuts on Public Desktop..." $statusTextBox

        # Tạo lối tắt cho Google Chrome
        $chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        if (Test-Path $chromePath) {
            $shortcutPath = Join-Path $publicDesktop "Google Chrome.lnk"
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($shortcutPath)
            $Shortcut.TargetPath = $chromePath
            $Shortcut.Save()
            Add-Status "Created shortcut for Google Chrome." $statusTextBox
        }

        # Tạo lối tắt cho Unikey
        $unikeyPath = "C:\unikey46RC2-230919-win64\UniKeyNT.exe"
        if (Test-Path $unikeyPath) {
            $shortcutPath = Join-Path $publicDesktop "Unikey.lnk"
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($shortcutPath)
            $Shortcut.TargetPath = $unikeyPath
            $Shortcut.Save()
            Add-Status "Created shortcut for Unikey." $statusTextBox
        }

        return $true
    }
    catch {
        Add-Status "ERROR during System Configuration: $_" $statusTextBox
        return $false
    }
}



# Function to handle Run All operations
function Invoke-RunAllOperations {
    param (
        [System.Windows.Forms.Form]$mainForm
    )
    
    # Create status form
    $statusForm = New-Object System.Windows.Forms.Form
    $statusForm.Text = "Running All Operations"
    $statusForm.Size = New-Object System.Drawing.Size(595, 500)
    $statusForm.StartPosition = "CenterScreen"
    $statusForm.BackColor = [System.Drawing.Color]::Black
    $statusForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $statusForm.MaximizeBox = $false
    $statusForm.MinimizeBox = $false

    # Add gradient background
    $statusForm.Add_Paint({
        $graphics = $_.Graphics
        $rect = New-Object System.Drawing.Rectangle(0, 0, $statusForm.Width, $statusForm.Height)
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $rect,
            [System.Drawing.Color]::FromArgb(0, 0, 0),
            [System.Drawing.Color]::FromArgb(0, 40, 0),
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
        )
        $graphics.FillRectangle($brush, $rect)
        $brush.Dispose()
    })

    # Title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "RUNNING ALL OPERATIONS"
    $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(580, 30)
    $titleLabel.ForeColor = [System.Drawing.Color]::Lime
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $statusForm.Controls.Add($titleLabel)

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(10, 60)
    $statusTextBox.Size = New-Object System.Drawing.Size(560, 350)
    $statusTextBox.BackColor = [System.Drawing.Color]::Black
    $statusTextBox.ForeColor = [System.Drawing.Color]::Lime
    $statusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $statusTextBox.ReadOnly = $true
    $statusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $statusForm.Controls.Add($statusTextBox)

    # Function to add status message
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

    # Progress bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 420)
    $progressBar.Size = New-Object System.Drawing.Size(560, 30)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $statusForm.Controls.Add($progressBar)

    # Show the form
    $statusForm.Show()
    [System.Windows.Forms.Application]::DoEvents()

    try {
        # Step 1: Device Selection and Software Installation
        Add-Status "STEP 1/7: Selecting Device Type and Installing Software..." $statusTextBox
        $progressBar.Value = 14

        # Create device selection form
        $deviceForm = New-Object System.Windows.Forms.Form
        $deviceForm.Text = "Select Device Type"
        $deviceForm.Size = New-Object System.Drawing.Size(300, 210)
        $deviceForm.StartPosition = "CenterScreen"
        $deviceForm.BackColor = [System.Drawing.Color]::Black
        $deviceForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $deviceForm.MaximizeBox = $false
        $deviceForm.MinimizeBox = $false
        $deviceForm.Add_Paint({
            $graphics = $_.Graphics
            $rect = New-Object System.Drawing.Rectangle(0, 0, $deviceForm.Width, $deviceForm.Height)
            $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
                $rect,
                [System.Drawing.Color]::FromArgb(0, 0, 0),
                [System.Drawing.Color]::FromArgb(0, 40, 0),
                [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
            )
            $graphics.FillRectangle($brush, $rect)
            $brush.Dispose()
        })

        # Title label
        $deviceTitleLabel = New-Object System.Windows.Forms.Label
        $deviceTitleLabel.Text = "SELECT DEVICE TYPE"
        $deviceTitleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $deviceTitleLabel.Size = New-Object System.Drawing.Size(290, 30)
        $deviceTitleLabel.ForeColor = [System.Drawing.Color]::Lime
        $deviceTitleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $deviceTitleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $deviceTitleLabel.BackColor = [System.Drawing.Color]::Transparent
        $deviceForm.Controls.Add($deviceTitleLabel)

        # Desktop button
        $btnDesktop = New-DynamicButton -text "DESKTOP" -x 10 -y 70 -width 260 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            $script:selectedDeviceType = "Desktop"
            $deviceForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $deviceForm.Close()
        }
        $deviceForm.Controls.Add($btnDesktop)

        # Laptop button
        $btnLaptop = New-DynamicButton -text "LAPTOP" -x 10 -y 120 -width 260 -height 40 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
            $script:selectedDeviceType = "Laptop"
            $deviceForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $deviceForm.Close()
        }
        $deviceForm.Controls.Add($btnLaptop)

        # Show device selection form and get result
        $result = $deviceForm.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $deviceType = $script:selectedDeviceType
            Add-Status "Selected device type: $deviceType" $statusTextBox
        } else {
            Add-Status "Device type selection cancelled. Exiting..." $statusTextBox
            return
        }

        # Copy software files (gọi hàm toàn cục)
        Add-Status "Copying software files..." $statusTextBox
        $copyResult = Copy-SoftwareFiles -deviceType $deviceType $statusTextBox
        if (-not $copyResult) {
            Add-Status "Error copying software files. Exiting..." $statusTextBox
            return
        }

        # Install software (gọi hàm toàn cục)
        Add-Status "Installing software..." $statusTextBox
        Install-Software -deviceType $deviceType $statusTextBox
        Add-Status "Software installation completed successfully for $deviceType" $statusTextBox
        Add-Status "STEP 1/7 completed successfully!" $statusTextBox

        # Step 2: System Configuration and Shortcut Creation
        Add-Status "Step 2/7: Configuring System and Creating Shortcuts..." $statusTextBox
        $progressBar.Value = 28 # Tăng giá trị progress bar

        $configResult = Invoke-SystemConfiguration -deviceType $deviceType -statusTextBox $statusTextBox
        if ($configResult) {
            Add-Status "Step 2 completed successfully!" $statusTextBox
        } else {
            Add-Status "Step 2 encountered errors. Check logs." $statusTextBox
        }
    }
    catch {
        Add-Status "Error occurred: $_" $statusTextBox
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred during the operations: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        # Close the status form after a delay
        # Start-Sleep -Seconds 2
        # $statusForm.Close()
    }
}

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
    $deviceTypeForm.Add_Paint({
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
    })

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

    # Desktop button
    $btnDesktop = New-DynamicButton -text "DESKTOP" -x 10 -y 60 -width 200 -height 50 -clickAction {
        Add-Status "STEP 1: Copying required files for Desktop..." $statusTextBox
        $copyResult = Copy-SoftwareFiles -deviceType "Desktop" $statusTextBox

        if ($copyResult) {
            Add-Status "STEP 2: Installing software for Desktop..." $statusTextBox
            $installResult = Install-Software -deviceType "Desktop" $statusTextBox

            if ($installResult) {
                Add-Status "All software installation completed successfully!" $statusTextBox
            }
            else {
                Add-Status "Warning: Some installations may have failed." $statusTextBox
            }
        }
        else {
            Add-Status "Error: Failed to copy required files. Installation aborted." $statusTextBox
        }
    }
    $deviceTypeForm.Controls.Add($btnDesktop)

    # Laptop button
    $btnLaptop = New-DynamicButton -text "LAPTOP" -x 260 -y 60 -width 200 -height 50 -clickAction {
        Add-Status "STEP 1: Copying required files for Laptop..." $statusTextBox
        $copyResult = Copy-SoftwareFiles -deviceType "Laptop" $statusTextBox

        if ($copyResult) {
            Add-Status "STEP 2: Installing software for Laptop..." $statusTextBox
            $installResult = Install-Software -deviceType "Laptop" $statusTextBox

            if ($installResult) {
                Add-Status "All software installation completed successfully!" $statusTextBox
            }
            else {
                Add-Status "Warning: Some installations may have failed." $statusTextBox
            }
        }
        else {
            Add-Status "Error: Failed to copy required files. Installation aborted." $statusTextBox
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
# [1] Run All Functions
$buttonRunAll = New-DynamicButton -text "[1] Run All" -x 30 -y 100 -width 380 -height 60 -clickAction {
    Invoke-RunAllOperations -mainForm $script:form
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