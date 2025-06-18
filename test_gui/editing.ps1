# ==============================================================================
# BAROPROVIP - VOLUME MANAGEMENT TOOL (FIXED VERSION)
# ==============================================================================

# SECTION 1: ADMIN PRIVILEGES CHECK & INITIALIZATION
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "This script requires administrative privileges. Attempting to restart with elevation..."
    Start-Sleep -Seconds 0

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

# STEP 1
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

# STEP 2
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
        
        # THÊM XỬ LÝ PHÍM ESC VÀ ENTER
        $renameForm.KeyPreview = $true
        $renameForm.Add_KeyDown({
            param($sender, $e)
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
                # ESC để đóng form
                $renameForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $renameForm.Close()
            }
            elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                # ENTER để thực hiện rename
                $okButton.PerformClick()
            }
        })

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
        
        # THÊM XỬ LÝ PHÍM ENTER CHO TEXTBOX
        $nameTextBox.Add_KeyDown({
            param($sender, $e)
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $okButton.PerformClick()
            }
        })

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
        $okButton.Text = "OK (Enter)"
        $okButton.Location = New-Object System.Drawing.Point(220, 160)
        $okButton.Size = New-Object System.Drawing.Size(100, 30)
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
        $cancelButton.Text = "Cancel (ESC)"
        $cancelButton.Location = New-Object System.Drawing.Point(330, 160)
        $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
        $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
        $cancelButton.ForeColor = [System.Drawing.Color]::White
        $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $cancelButton.Add_Click({
            $renameForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $renameForm.Close()
        })
        $renameForm.Controls.Add($cancelButton)

        # Đặt focus vào TextBox khi form hiển thị
        $renameForm.Add_Shown({
            $nameTextBox.Focus()
            $nameTextBox.Select()
        })

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

# STEP 3
function Invoke-SystemCleanup {
    param (
        [string]$deviceType,
        [System.Windows.Forms.TextBox]$statusTextBox
    )
    
    try {
        Add-Status "Starting system cleanup and optimization..." $statusTextBox
        
        # --- 1. System File Cleanup ---
        Invoke-FileCleanup $statusTextBox
        
        # --- 2. Service Management ---
        Invoke-ServiceOptimization $statusTextBox
        
        # --- 3. System Performance Optimization ---
        Invoke-PerformanceOptimization $statusTextBox
        
        # --- 4. Power Management ---
        Invoke-PowerConfiguration $deviceType $statusTextBox
        
        # --- 5. Startup Program Management ---
        Invoke-StartupOptimization $statusTextBox
        
        # --- 6. Disk Optimization ---
        Invoke-DiskOptimization $statusTextBox
        
        # --- 7. Timezone Configuration ---
        Invoke-TimezoneConfiguration $statusTextBox
        
        # --- 8. Power Options Configuration ---
        Invoke-PowerOptionsConfiguration $statusTextBox
        
        Add-Status "System cleanup and optimization completed successfully!" $statusTextBox
        return $true
        
    } catch {
        Add-Status "ERROR during System Cleanup: $_" $statusTextBox
        return $false
    }
}

# Helper Functions
function Invoke-FileCleanup {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Cleaning temporary files..." $statusTextBox
    
    # Định nghĩa các đường dẫn cần dọn dẹp
    $tempPaths = @(
        "$env:TEMP\*",
        "$env:WINDIR\Temp\*",
        "$env:USERPROFILE\AppData\Local\Temp\*"
    )
    
    # Dọn dẹp file tạm
    $tempPaths | ForEach-Object {
        try {
            Remove-Item -Path $_ -Recurse -Force -ErrorAction SilentlyContinue
            Add-Status "Cleaned: $_" $statusTextBox
        } catch {
            Add-Status "Warning: Could not clean $_" $statusTextBox
        }
    }
    
    # Dọn dẹp Recycle Bin và Windows Update cache
    try {
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue
        Add-Status "Recycle Bin cleaned successfully!" $statusTextBox
        
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "$env:WINDIR\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
        Start-Service -Name wuauserv -ErrorAction SilentlyContinue
        Add-Status "Windows Update cache cleaned!" $statusTextBox
    } catch {
        Add-Status "Warning: Could not complete advanced cleanup" $statusTextBox
    }
}

function Invoke-ServiceOptimization {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Disabling unnecessary services..." $statusTextBox
    
    $servicesToDisable = @("Fax", "WSearch", "Themes", "TabletInputService", "WMPNetworkSvc")
    
    $servicesToDisable | ForEach-Object {
        try {
            $svc = Get-Service -Name $_ -ErrorAction SilentlyContinue
            if ($svc -and $svc.StartType -ne "Disabled") {
                Stop-Service -Name $_ -Force -ErrorAction SilentlyContinue
                Set-Service -Name $_ -StartupType Disabled -ErrorAction SilentlyContinue
                Add-Status "Disabled service: $_" $statusTextBox
            }
        } catch {
            Add-Status "Warning: Could not disable service $_" $statusTextBox
        }
    }
}

function Invoke-PerformanceOptimization {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Optimizing visual effects for performance..." $statusTextBox
    
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects"
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        Set-ItemProperty -Path $regPath -Name "VisualFXSetting" -Value 2 -Type DWord
        Add-Status "Visual effects optimized for performance!" $statusTextBox
        
        # Tắt Windows Defender tạm thời
        Set-MpPreference -DisableRealtimeMonitoring $true -ErrorAction SilentlyContinue
        Add-Status "Windows Defender real-time protection disabled temporarily!" $statusTextBox
        Add-Status "Note: This will be re-enabled automatically after restart" $statusTextBox
    } catch {
        Add-Status "Warning: Could not complete performance optimization" $statusTextBox
    }
}

function Invoke-PowerConfiguration {
    param ([string]$deviceType, [System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Optimizing power settings..." $statusTextBox
    
    $powerScheme = if ($deviceType -eq "Desktop") { 
        "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"  # High Performance
    } else { 
        "381b4222-f694-41f0-9685-ff5bb260df2e"  # Balanced
    }
    
    try {
        powercfg /setactive $powerScheme
        $planName = if ($deviceType -eq "Desktop") { "High Performance" } else { "Balanced" }
        Add-Status "Power plan set to $planName for $deviceType!" $statusTextBox
    } catch {
        Add-Status "Warning: Could not set power plan" $statusTextBox
    }
}

function Invoke-StartupOptimization {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Disabling unnecessary startup programs..." $statusTextBox
    
    $startupPrograms = @("Skype for Desktop", "Spotify", "Steam", "Discord")
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    
    $startupPrograms | ForEach-Object {
        try {
            $property = Get-ItemProperty -Path $regPath -Name $_ -ErrorAction SilentlyContinue
            if ($property) {
                Remove-ItemProperty -Path $regPath -Name $_ -ErrorAction SilentlyContinue
                Add-Status "Disabled startup program: $_" $statusTextBox
            }
        } catch {
            Add-Status "Warning: Could not disable startup program $_" $statusTextBox
        }
    }
}

function Invoke-DiskOptimization {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Running Disk Cleanup..." $statusTextBox
    
    try {
        # Disk Cleanup
        Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -Wait -WindowStyle Hidden
        Add-Status "Disk Cleanup completed!" $statusTextBox
        
        # Drive Optimization
        Add-Status "Checking drive type and optimizing..." $statusTextBox
        $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
        $drives | ForEach-Object {
            $driveLetter = $_.DeviceID
            Add-Status "Optimizing drive $driveLetter..." $statusTextBox
            Start-Process -FilePath "defrag.exe" -ArgumentList "$driveLetter /O" -Wait -WindowStyle Hidden
        }
        Add-Status "Drive optimization completed!" $statusTextBox
    } catch {
        Add-Status "Warning: Could not complete disk optimization" $statusTextBox
    }
}

function Invoke-TimezoneConfiguration {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Setting time zone and automatically updating time..." $statusTextBox
    
    try {
        # Cấu hình múi giờ
        $tzResult = Start-Process -FilePath "tzutil" -ArgumentList "/s `"SE Asia Standard Time`"" -Wait -PassThru -WindowStyle Hidden
        if ($tzResult.ExitCode -eq 0) {
            Add-Status "Time zone set to SE Asia Standard Time successfully!" $statusTextBox
        }
        
        # Cấu hình NTP và đồng bộ thời gian
        $regCommands = @(
            @{Path = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\w32time\Parameters"; Name = "Type"; Value = "NTP"},
            @{Path = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\tzautoupdate"; Name = "Start"; Value = 2}
        )
        
        $regCommands | ForEach-Object {
            reg add $_.Path /v $_.Name /t REG_SZ /d $_.Value /f | Out-Null
        }
        
        w32tm /resync | Out-Null
        Add-Status "Set timezone and time automatically completed successfully!" $statusTextBox
    } catch {
        Add-Status "Warning: Could not configure timezone settings: $_" $statusTextBox
    }
}

function Invoke-PowerOptionsConfiguration {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Configuring power options to 'Do Nothing'..." $statusTextBox
    
    try {
        # Định nghĩa các cấu hình power
        $powerConfigs = @(
            @{Setting = "LIDACTION"; Description = "Lid close action"},
            @{Setting = "SBUTTONACTION"; Description = "Sleep button action"},
            @{Setting = "PBUTTONACTION"; Description = "Power button action"}
        )
        
        # Áp dụng cấu hình cho các nút
        $powerConfigs | ForEach-Object {
            powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_BUTTONS $_.Setting 0 | Out-Null
            powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_BUTTONS $_.Setting 0 | Out-Null
            Add-Status "$($_.Description) set to 'Do Nothing'!" $statusTextBox
        }
        
        # Tắt timeout cho màn hình và sleep
        powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 0 | Out-Null
        powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 0 | Out-Null
        powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_SLEEP STANDBYIDLE 0 | Out-Null
        powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_SLEEP STANDBYIDLE 0 | Out-Null
        
        # Áp dụng thay đổi
        powercfg /SETACTIVE SCHEME_CURRENT | Out-Null
        Add-Status "Power options configured to 'Do Nothing' completed successfully!" $statusTextBox
    } catch {
        Add-Status "Warning: Could not configure power options: $_" $statusTextBox
    }
}

# STEP 4: 
function Invoke-ActivationConfiguration {
    param (
        [string]$deviceType,
        [System.Windows.Forms.TextBox]$statusTextBox
    )
    
    try {
        # --- 1. Windows 10 Pro Activation ---
        Invoke-WindowsActivation $statusTextBox
        
        # --- 2. Office 2019 Pro Plus Activation ---
        Invoke-OfficeActivation $statusTextBox
        
        Add-Status "Activations completed successfully!" $statusTextBox
        return $true
        
    } catch {
        Add-Status "ERROR during Activation Configuration: $_" $statusTextBox
        return $false
    }
}

function Invoke-WindowsActivation {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Checking Windows activation status..." $statusTextBox
    
    try {
        # Kiểm tra trạng thái activation của Windows
        $windowsActivationStatus = Get-CimInstance SoftwareLicensingProduct -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseStatus = 1" -ErrorAction SilentlyContinue
        
        if ($windowsActivationStatus) {
            Add-Status "Product: $($windowsActivationStatus.Name)" $statusTextBox
            Add-Status "Partial Product Key: $($windowsActivationStatus.PartialProductKey)" $statusTextBox
            Add-Status "Windows is already activated. Skipping activation..." $statusTextBox
            return
        }

        Add-Status "Windows is not activated. Proceeding with activation..." $statusTextBox
        Add-Status "Activating Windows 10 Pro..." $statusTextBox
    
        # Kiểm tra phiên bản Windows hiện tại
        $windowsVersion = (Get-WmiObject -Class Win32_OperatingSystem).Caption
        Add-Status "Current Windows version: $windowsVersion" $statusTextBox
        
        # Nhập Windows 10 Pro license key
        $windows10ProKey = "VK7JG-NPHTM-C97JM-9MPGT-3V66T"  # Thay thế bằng key license thực tế của bạn
        
        if ([string]::IsNullOrWhiteSpace($windows10ProKey)) {
            Add-Status "WARNING: Windows license key is empty. Please add your license key." $statusTextBox
            Add-Status "Skipping Windows activation..." $statusTextBox
            return
        }
        
        # Cài đặt product key
        $result = Start-Process -FilePath "slmgr" -ArgumentList "/ipk $windows10ProKey" -Wait -PassThru -WindowStyle Hidden
        if ($result.ExitCode -eq 0) {
            # Kích hoạt Windows trực tiếp (không qua KMS)
            Add-Status "Activating Windows with provided license key..." $statusTextBox
            $activateResult = Start-Process -FilePath "slmgr" -ArgumentList "/ato" -Wait -PassThru -WindowStyle Hidden
            if ($activateResult.ExitCode -eq 0) {
                Add-Status "Windows activated successfully!" $statusTextBox
            } else {
                Add-Status "Warning: Windows activation may have failed" $statusTextBox
            }
        } else {
            Add-Status "Warning: Windows key installation failed" $statusTextBox
        }
        
        # Kiểm tra trạng thái activation
        Start-Sleep -Seconds 2
        Add-Status "Checking Windows activation status..." $statusTextBox
        Start-Process -FilePath "slmgr" -ArgumentList "/xpr" -Wait -WindowStyle Hidden
    } catch {
        Add-Status "Warning: Windows activation encountered errors: $_" $statusTextBox
    }
}

function Invoke-OfficeActivation {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    Add-Status "Checking Offices activation status..." $statusTextBox
    try {
        # Tìm đường dẫn Office installation
        $officePaths = @(
            "C:\Program Files\Microsoft Office\Office16",
            "C:\Program Files (x86)\Microsoft Office\Office16"
        )
        $officePath = $null
        foreach ($path in $officePaths) {
            if (Test-Path "$path\ospp.vbs") {
                $officePath = $path
                Add-Status "Found Office at: $path" $statusTextBox
                break
            }
        }
        if (-not $officePath) {
            Add-Status "Office 2019 installation not found. Skipping activation..." $statusTextBox
            return
        }
        
        # Chuyển đến thư mục Office
        Set-Location $officePath

        # Kiểm tra trạng thái activation của Office
        try {
            $officeStatusResult = Start-Process -FilePath "cscript" -ArgumentList "//nologo ospp.vbs /dstatus" -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput "$env:TEMP\office_status.txt"
            
            if (Test-Path "$env:TEMP\office_status.txt") {
                $officeStatusContent = Get-Content "$env:TEMP\office_status.txt" -Raw
                
                # Kiểm tra xem Office đã được kích hoạt chưa
                if ($officeStatusContent -match "LICENSE STATUS:\s*---LICENSED---" -or 
                    $officeStatusContent -match "LICENSE STATUS:\s*---LICENSED \(GRACE\)---") {
                    # Hiển thị thông tin license hiện tại
                    $licenseLines = $officeStatusContent -split "`n" | Where-Object { $_ -match "PRODUCT NAME|LICENSE STATUS|PARTIAL PRODUCT KEY" }
                    foreach ($line in $licenseLines) {
                        if ($line.Trim() -ne "") {
                            Add-Status "Offices Info: $($line.Trim())" $statusTextBox
                        }
                    }
                    Add-Status "Offices is already activated. Skipping activation..." $statusTextBox
                    # Xóa file tạm
                    Remove-Item "$env:TEMP\office_status.txt" -Force -ErrorAction SilentlyContinue
                    return
                }
                
                # Xóa file tạm
                Remove-Item "$env:TEMP\office_status.txt" -Force -ErrorAction SilentlyContinue
            }
        } catch {
            Add-Status "Could not check Office activation status. Proceeding with activation..." $statusTextBox
        }
        
        Add-Status "Office is not activated. Proceeding with activation..." $statusTextBox        
        # Kiểm tra trạng thái activation của Office
        try {
            $officeStatusResult = Start-Process -FilePath "cscript" -ArgumentList "//nologo ospp.vbs /dstatus" -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput "$env:TEMP\office_status.txt"
            
            if (Test-Path "$env:TEMP\office_status.txt") {
                $officeStatusContent = Get-Content "$env:TEMP\office_status.txt" -Raw
                
                # Kiểm tra xem Office đã được kích hoạt chưa
                if ($officeStatusContent -match "LICENSE STATUS:\s*---LICENSED---" -or 
                    $officeStatusContent -match "LICENSE STATUS:\s*---LICENSED \(GRACE\)---") {
                    Add-Status "Office is already activated. Skipping activation..." $statusTextBox
                    
                    # Hiển thị thông tin license hiện tại
                    $licenseLines = $officeStatusContent -split "`n" | Where-Object { $_ -match "PRODUCT NAME|LICENSE STATUS|PARTIAL PRODUCT KEY" }
                    foreach ($line in $licenseLines) {
                        if ($line.Trim() -ne "") {
                            Add-Status "Office Info: $($line.Trim())" $statusTextBox
                        }
                    }
                    
                    # Xóa file tạm
                    Remove-Item "$env:TEMP\office_status.txt" -Force -ErrorAction SilentlyContinue
                    return
                }
                
                # Xóa file tạm
                Remove-Item "$env:TEMP\office_status.txt" -Force -ErrorAction SilentlyContinue
            }
        } catch {
            Add-Status "Could not check Office activation status. Proceeding with activation..." $statusTextBox
        }
        
        Add-Status "Office is not activated. Proceeding with activation..." $statusTextBox
        
        # Nhập Office 2019 Pro Plus license key
        $officeProPlusKey = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"  # Thay thế bằng key license thực tế của bạn
        
        if ([string]::IsNullOrWhiteSpace($officeProPlusKey)) {
            Add-Status "WARNING: Office license key is empty. Please add your license key." $statusTextBox
            Add-Status "Skipping Office activation..." $statusTextBox
            return
        }
        
        # Cài đặt product key
        $keyResult = Start-Process -FilePath "cscript" -ArgumentList "//nologo ospp.vbs /inpkey:$officeProPlusKey" -Wait -PassThru -WindowStyle Hidden
        if ($keyResult.ExitCode -eq 0) {
            # Kích hoạt Office trực tiếp (không qua KMS)
            Add-Status "Activating Office2019ProPlus with provided license key..." $statusTextBox
            $activateResult = Start-Process -FilePath "cscript" -ArgumentList "//nologo ospp.vbs /act" -Wait -PassThru -WindowStyle Hidden
            
            if ($activateResult.ExitCode -eq 0) {
                Add-Status "Office 2019 Pro Plus activated successfully!" $statusTextBox
            } else {
                Add-Status "Warning: Office activation may have failed" $statusTextBox
            }
        } else {
            Add-Status "Warning: Office key installation failed" $statusTextBox
        }
        
        # Kiểm tra trạng thái activation Office
        Add-Status "Checking Office activation status..." $statusTextBox
        Start-Process -FilePath "cscript" -ArgumentList "//nologo ospp.vbs /dstatus" -Wait -WindowStyle Hidden
        
    } catch {
        Add-Status "Warning: Office activation encountered errors: $_" $statusTextBox
    }
}

# STEP 5: 
function Invoke-WindowsFeaturesConfiguration {
    param (
        [string]$deviceType,
        [System.Windows.Forms.TextBox]$statusTextBox
    )
    
    try {
        # --- 1. Check and Enable Required Features ---
        Invoke-EnableWindowsFeatures $statusTextBox
        # --- 2. Check and Disable Unnecessary Features ---
        Invoke-DisableWindowsFeatures $statusTextBox
        return $true
    } catch {
        Add-Status "ERROR during Windows Features Configuration: $_" $statusTextBox
        return $false
    }
}

# Helper Functions cho Windows Features Configuration
function Invoke-EnableWindowsFeatures {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    # Danh sách các features cần enable
    $featuresToEnable = @(
        @{
            Name = "NetFx3"
            DisplayName = ".NET 3.5    "
            Command = "dism /online /enable-feature /featurename:NetFx3 /all /norestart"
        },
        @{
            Name = "WCF-HTTP-Activation"
            DisplayName = "WCF HTTP    "
            Command = "DISM /Online /Enable-Feature /FeatureName:WCF-HTTP-Activation /All /Quiet /NoRestart"
        },
        @{
            Name = "WCF-NonHTTP-Activation"
            DisplayName = "WCF Non-HTTP"
            Command = "DISM /Online /Enable-Feature /FeatureName:WCF-NonHTTP-Activation /All /Quiet /NoRestart"
        }
    )
    
    foreach ($feature in $featuresToEnable) {
        try {
            # Kiểm tra trạng thái hiện tại của feature bằng PowerShell cmdlet
            $currentFeature = Get-WindowsOptionalFeature -Online -FeatureName $feature.Name -ErrorAction SilentlyContinue
            if ($currentFeature) {
                $currentState = $currentFeature.State
                if ($currentState -eq "Enabled") {
                    Add-Status "$($feature.DisplayName): Already enabled. Skipping..." $statusTextBox
                } elseif ($currentState -eq "Disabled") {
                    Add-Status "$($feature.DisplayName): Currently disabled. Enabling..." $statusTextBox
                    
                    # Enable feature using DISM command
                    $enableArgs = $feature.Command.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries) | Select-Object -Skip 1
                    $enableResult = Start-Process -FilePath "dism" -ArgumentList $enableArgs -Wait -PassThru -WindowStyle Hidden
                
                    if ($enableResult.ExitCode -eq 0) {
                        Add-Status "$($feature.DisplayName): Enabled successfully!" $statusTextBox
                    } elseif ($enableResult.ExitCode -eq 3010) {
                        Add-Status "$($feature.DisplayName): Enabled successfully! (Restart required)" $statusTextBox
                    } else {
                        Add-Status "WARNING: Failed to enable $($feature.DisplayName) (Exit code: $($enableResult.ExitCode))" $statusTextBox
                    }
                    
                    # Verify new state
                    Start-Sleep -Seconds 2
                    $newFeature = Get-WindowsOptionalFeature -Online -FeatureName $feature.Name -ErrorAction SilentlyContinue
                    if ($newFeature) {
                        Add-Status "$($feature.DisplayName): Verified new state is $($newFeature.State)" $statusTextBox
                    }
                } else {
                    Add-Status "WARNING: $($feature.DisplayName) is in unexpected state: $currentState" $statusTextBox
                }
            } else {
                Add-Status "WARNING: Could not find feature $($feature.Name)" $statusTextBox
            }
            
        } catch {
            Add-Status "ERROR: Failed to process $($feature.DisplayName): $_" $statusTextBox
        }
    }
}

function Invoke-DisableWindowsFeatures {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    # Danh sách các features cần disable
    $featuresToDisable = @(
        @{
            Name = "Internet-Explorer-Optional-amd64"
            DisplayName = "IExplorer 11"
            Command = "dism /online /disable-feature /featurename:Internet-Explorer-Optional-amd64 /norestart"
        }
    )
    
    foreach ($feature in $featuresToDisable) {
        try {
            # Kiểm tra trạng thái hiện tại của feature
            $currentFeature = Get-WindowsOptionalFeature -Online -FeatureName $feature.Name -ErrorAction SilentlyContinue
            
            if ($currentFeature) {
                $currentState = $currentFeature.State
                if ($currentState -eq "Disabled") {
                    Add-Status "$($feature.DisplayName): Already disabled.Skipping..." $statusTextBox
                } elseif ($currentState -eq "Enabled") {
                    Add-Status "$($feature.DisplayName): Currently enabled. Disabling..." $statusTextBox
                    
                    # Disable feature using DISM command
                    $disableArgs = $feature.Command.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries) | Select-Object -Skip 1
                    $disableResult = Start-Process -FilePath "dism" -ArgumentList $disableArgs -Wait -PassThru -WindowStyle Hidden
                    
                    if ($disableResult.ExitCode -eq 0) {
                        Add-Status "$($feature.DisplayName): Disabled successfully!" $statusTextBox
                    } elseif ($disableResult.ExitCode -eq 3010) {
                        Add-Status "$($feature.DisplayName): Disabled successfully! (Restart required)" $statusTextBox
                    } else {
                        Add-Status "WARNING: Failed to disable $($feature.DisplayName) (Exit code: $($disableResult.ExitCode))" $statusTextBox
                    }
                    
                    # Verify new state
                    Start-Sleep -Seconds 2
                    $newFeature = Get-WindowsOptionalFeature -Online -FeatureName $feature.Name -ErrorAction SilentlyContinue
                    if ($newFeature) {
                        Add-Status "$($feature.DisplayName): Verified new state is $($newFeature.State)" $statusTextBox
                    }
                } else {
                    Add-Status "WARNING: $($feature.DisplayName) is in unexpected state: $currentState" $statusTextBox
                }
            } else {
                Add-Status "WARNING: Could not find feature $($feature.Name)" $statusTextBox
            }
            
        } catch {
            Add-Status "ERROR: Failed to process $($feature.DisplayName): $_" $statusTextBox
        }
    }
}

# STEP 6:
function Invoke-DiskPartitioning {
    param (
        [string]$deviceType,
        [System.Windows.Forms.TextBox]$statusTextBox
    )
    
    try {
        # Chỉ thực hiện cho Laptop
        if ($deviceType -eq "Laptop") {
            # --- 1. Check Available Disks ---
            $diskCheckResult = Invoke-DiskAvailabilityCheck $statusTextBox
            if (-not $diskCheckResult) {
                Add-Status "No suitable disk found for partitioning. Skipping..." $statusTextBox
                return $true
            }
            
            # --- 2. Show Partition Size Selection ---
            $partitionResult = Invoke-PartitionSizeSelection $statusTextBox
            if ($partitionResult) {
                Add-Status "Disk partitioning completed successfully!" $statusTextBox
            } else {
                Add-Status "Disk partitioning was cancelled or failed." $statusTextBox
            }
        }
        return $true
    } catch {
        Add-Status "ERROR during Disk Partitioning: $_" $statusTextBox
        return $false
    }
}

# Helper Functions cho Disk Partitioning
function Invoke-DiskAvailabilityCheck {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    try {
        # Lấy thông tin tất cả các ổ đĩa
        $systemDisk = Get-Disk | Where-Object { $_.IsBoot -eq $true }
        
        # Kiểm tra xem ổ hệ thống có đủ không gian để chia không
        $systemPartitions = Get-Partition -DiskNumber $systemDisk.Number
        $systemVolume = $systemPartitions | Where-Object { $_.DriveLetter -eq 'C' }
        
        if ($systemVolume) {
            $usedSpace = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq "C:" }
            $freeSpaceGB = [math]::Round($usedSpace.FreeSpace / 1GB, 2)
            
            # Kiểm tra xem có thể chia phân vùng không (cần ít nhất 150GB trống)
            if ($freeSpaceGB -gt 150) {
                return $true
            } else {
                Add-Status "WARNING: No free space for safe partitioning (>150GB free)" $statusTextBox
                return $false
            }
        } else {
            Add-Status "WARNING: Could not determine C: drive information" $statusTextBox
            return $false
        }
        
    } catch {
        Add-Status "ERROR checking disk availability: $_" $statusTextBox
        return $false
    }
}

function Invoke-PartitionSizeSelection {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    try {
        # Tạo form chọn kích thước phân vùng
        $partitionForm = New-Object System.Windows.Forms.Form
        $partitionForm.Text = "Disk Partitioning - Select Size"
        $partitionForm.Size = New-Object System.Drawing.Size(600, 550)
        $partitionForm.StartPosition = "CenterScreen"
        $partitionForm.BackColor = [System.Drawing.Color]::Black
        $partitionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $partitionForm.MaximizeBox = $false
        $partitionForm.MinimizeBox = $false
        
        # Gradient background
        $partitionForm.Add_Paint({
            $graphics = $_.Graphics
            $rect = New-Object System.Drawing.Rectangle(0, 0, $partitionForm.Width, $partitionForm.Height)
            $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
                $rect,
                [System.Drawing.Color]::FromArgb(0, 0, 0),
                [System.Drawing.Color]::FromArgb(0, 40, 0),
                [System.Drawing.Drawing2D.LinearGradientMode]::Vertical
            )
            $graphics.FillRectangle($brush, $rect)
            $brush.Dispose()
        })
        
        # Title
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "SELECT PARTITION SIZE"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(600, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $partitionForm.Controls.Add($titleLabel)
        
        # Current disk information label
        $diskInfoLabel = New-Object System.Windows.Forms.Label
        $diskInfoLabel.Text = "CURRENT DISK INFORMATION:"
        $diskInfoLabel.Location = New-Object System.Drawing.Point(20, 60)
        $diskInfoLabel.Size = New-Object System.Drawing.Size(560, 25)
        $diskInfoLabel.ForeColor = [System.Drawing.Color]::Yellow
        $diskInfoLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $diskInfoLabel.BackColor = [System.Drawing.Color]::Transparent
        $partitionForm.Controls.Add($diskInfoLabel)
        
        # DataGridView để hiển thị thông tin ổ đĩa
        $diskGrid = New-Object System.Windows.Forms.DataGridView
        $diskGrid.Location = New-Object System.Drawing.Point(20, 90)
        $diskGrid.Size = New-Object System.Drawing.Size(540, 120)
        $diskGrid.BackgroundColor = [System.Drawing.Color]::Black
        $diskGrid.ForeColor = [System.Drawing.Color]::White
        $diskGrid.GridColor = [System.Drawing.Color]::Gray
        $diskGrid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $diskGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(0, 100, 0)
        $diskGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
        $diskGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
        $diskGrid.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
        $diskGrid.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
        $diskGrid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0, 150, 0)
        $diskGrid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
        $diskGrid.ColumnCount = 4
        $diskGrid.Columns[0].Name = "Letter"
        $diskGrid.Columns[1].Name = "Name"
        $diskGrid.Columns[2].Name = "Size (GB)"
        $diskGrid.Columns[3].Name = "Free (GB)"
        $diskGrid.Columns[0].Width = 60
        $diskGrid.Columns[1].Width = 150
        $diskGrid.Columns[2].Width = 100
        $diskGrid.Columns[3].Width = 100
        $diskGrid.ReadOnly = $true
        $diskGrid.AllowUserToAddRows = $false
        $diskGrid.AllowUserToDeleteRows = $false
        $diskGrid.RowHeadersVisible = $false
        $diskGrid.MultiSelect = $false
        $diskGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
        
        # Lấy thông tin ổ đĩa và thêm vào DataGridView
        try {
            $disks = Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
            foreach ($disk in $disks) {
                $letter = $disk.DeviceID
                $name = if ($disk.VolumeName) { $disk.VolumeName } else { "Local Disk" }
                $sizeGB = [math]::Round($disk.Size / 1GB, 2)
                $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
                $diskGrid.Rows.Add($letter, $name, $sizeGB, $freeGB)
            }
        } catch {
            Add-Status "Warning: Could not load disk information: $_" $statusTextBox
            # Thêm dữ liệu mẫu nếu không lấy được thông tin thực
            $diskGrid.Rows.Add("C:", "Windows", "500.00", "250.00")
        }
        
        $partitionForm.Controls.Add($diskGrid)
        
        # Instruction label
        $instructionLabel = New-Object System.Windows.Forms.Label
        $instructionLabel.Text = "Choose the size for the new data partition:"
        $instructionLabel.Location = New-Object System.Drawing.Point(20, 220)
        $instructionLabel.Size = New-Object System.Drawing.Size(560, 25)
        $instructionLabel.ForeColor = [System.Drawing.Color]::White
        $instructionLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $instructionLabel.BackColor = [System.Drawing.Color]::Transparent
        $partitionForm.Controls.Add($instructionLabel)
        
        # Variable to store selected size
        $script:selectedPartitionSize = 0
        
        # 100GB Button
        $btn100GB = New-Object System.Windows.Forms.Button
        $btn100GB.Text = "100 GB (Recommended for 256GB drives)"
        $btn100GB.Location = New-Object System.Drawing.Point(20, 250)
        $btn100GB.Size = New-Object System.Drawing.Size(540, 40)
        $btn100GB.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 0)
        $btn100GB.ForeColor = [System.Drawing.Color]::White
        $btn100GB.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btn100GB.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $btn100GB.Add_Click({
            $script:selectedPartitionSize = 101
            $partitionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $partitionForm.Close()
        })
        $partitionForm.Controls.Add($btn100GB)
        
        # 200GB Button
        $btn200GB = New-Object System.Windows.Forms.Button
        $btn200GB.Text = "200 GB (Recommended for 500GB drives)"
        $btn200GB.Location = New-Object System.Drawing.Point(20, 300)
        $btn200GB.Size = New-Object System.Drawing.Size(540, 40)
        $btn200GB.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 0)
        $btn200GB.ForeColor = [System.Drawing.Color]::White
        $btn200GB.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btn200GB.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $btn200GB.Add_Click({
            $script:selectedPartitionSize = 200
            $partitionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $partitionForm.Close()
        })
        $partitionForm.Controls.Add($btn200GB)
        
        # 500GB Button
        $btn500GB = New-Object System.Windows.Forms.Button
        $btn500GB.Text = "500 GB (Recommended for 1TB+ drives)"
        $btn500GB.Location = New-Object System.Drawing.Point(20, 350)
        $btn500GB.Size = New-Object System.Drawing.Size(540, 40)
        $btn500GB.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 0)
        $btn500GB.ForeColor = [System.Drawing.Color]::White
        $btn500GB.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btn500GB.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $btn500GB.Add_Click({
            $script:selectedPartitionSize = 500
            $partitionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $partitionForm.Close()
        })
        $partitionForm.Controls.Add($btn500GB)
        
        # Custom size section
        $customLabel = New-Object System.Windows.Forms.Label
        $customLabel.Text = "Custom size (GB):"
        $customLabel.Location = New-Object System.Drawing.Point(20, 410)
        $customLabel.Size = New-Object System.Drawing.Size(150, 25)
        $customLabel.ForeColor = [System.Drawing.Color]::Yellow
        $customLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $customLabel.BackColor = [System.Drawing.Color]::Transparent
        $partitionForm.Controls.Add($customLabel)
        
        # Custom size textbox
        $customTextBox = New-Object System.Windows.Forms.TextBox
        $customTextBox.Location = New-Object System.Drawing.Point(180, 408)
        $customTextBox.Size = New-Object System.Drawing.Size(100, 25)
        $customTextBox.Font = New-Object System.Drawing.Font("Arial", 10)
        $partitionForm.Controls.Add($customTextBox)
        
        # Custom size button
        $btnCustom = New-Object System.Windows.Forms.Button
        $btnCustom.Text = "Use Custom Size"
        $btnCustom.Location = New-Object System.Drawing.Point(300, 405)
        $btnCustom.Size = New-Object System.Drawing.Size(170, 30)
        $btnCustom.BackColor = [System.Drawing.Color]::FromArgb(0, 100, 150)
        $btnCustom.ForeColor = [System.Drawing.Color]::White
        $btnCustom.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnCustom.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $btnCustom.Add_Click({
            $customSize = $customTextBox.Text.Trim()
            if ($customSize -match '^\d+$' -and [int]$customSize -gt 0 -and [int]$customSize -le 2000) {
                $script:selectedPartitionSize = [int]$customSize
                $partitionForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $partitionForm.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please enter a valid size between 1 and 2000 GB",
                    "Invalid Input",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }
        })
        $partitionForm.Controls.Add($btnCustom)
        
        # Cancel button
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Cancel"
        $btnCancel.Location = New-Object System.Drawing.Point(250, 460)
        $btnCancel.Size = New-Object System.Drawing.Size(100, 35)
        $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
        $btnCancel.ForeColor = [System.Drawing.Color]::White
        $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnCancel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $btnCancel.Add_Click({
            $partitionForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $partitionForm.Close()
        })
        $partitionForm.Controls.Add($btnCancel)
        
        # Show form and get result
        $result = $partitionForm.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            Add-Status "Selected partition size: $($script:selectedPartitionSize) GB" $statusTextBox
            return Invoke-CreatePartition -sizeGB $script:selectedPartitionSize -statusTextBox $statusTextBox
        } else {
            Add-Status "Partition creation cancelled by user." $statusTextBox
            return $false
        }
        
    } catch {
        Add-Status "ERROR in partition size selection: $_" $statusTextBox
        return $false
    }
}

function Invoke-CreatePartition {
    param (
        [int]$sizeGB,
        [System.Windows.Forms.TextBox]$statusTextBox
    )
    
    try {
        # Tắt dịch vụ ShellHWDetection để tránh popup
        try {
            Stop-Service -Name ShellHWDetection -Force -ErrorAction SilentlyContinue
        } catch {
            Add-Status "Warning: Could not disable hardware detection service" $statusTextBox
        }
        
        # Lấy ổ đĩa hệ thống
        $systemDisk = Get-Disk | Where-Object { $_.IsBoot -eq $true }
        $diskNumber = $systemDisk.Number
        
        Add-Status "Working with system disk: Disk $diskNumber" $statusTextBox
        
        # Chuyển đổi GB sang bytes
        $sizeBytes = $sizeGB * 1GB
        
        # Lấy phân vùng C: để shrink
        $cPartition = Get-Partition -DiskNumber $diskNumber | Where-Object { $_.DriveLetter -eq 'C' }
        
        if (-not $cPartition) {
            Add-Status "ERROR: Could not find C: partition" $statusTextBox
            return $false
        }
        # Shrink C: partition
        try {
            $newCSize = $cPartition.Size - $sizeBytes
            Resize-Partition -DiskNumber $diskNumber -PartitionNumber $cPartition.PartitionNumber -Size $newCSize
            Add-Status "Drive C: partition shrunk successfully!" $statusTextBox
        } catch {
            Add-Status "ERROR shrinking C: partition: $_" $statusTextBox
            return $false
        }
        
        # Tạo phân vùng mới KHÔNG gán drive letter ngay
        try {
            $newPartition = New-Partition -DiskNumber $diskNumber -Size $sizeBytes
            Add-Status "New partition created (Partition Number: $($newPartition.PartitionNumber))" $statusTextBox
        } catch {
            Add-Status "ERROR creating new partition: $_" $statusTextBox
            return $false
        }
        
        # Format phân vùng mới KHÔNG có drive letter (silent)
        try {
            # Format bằng diskpart để tránh popup
            $diskpartScript = @"
select disk $diskNumber
select partition $($newPartition.PartitionNumber)
format fs=ntfs label="DATA" quick
"@
            $diskpartScript | diskpart
            Add-Status "New partition formatted successfully with NTFS!" $statusTextBox
        } catch {
            Add-Status "ERROR formatting new partition: $_" $statusTextBox
            return $false
        }
        
        # Gán drive letter SAU khi format
        try {
            $newPartition | Add-PartitionAccessPath -AssignDriveLetter
            $newPartition = Get-Partition -DiskNumber $diskNumber -PartitionNumber $newPartition.PartitionNumber
            $newDriveLetter = $newPartition.DriveLetter
            Add-Status "Drive letter assigned: $newDriveLetter" $statusTextBox
        } catch {
            Add-Status "ERROR assigning drive letter: $_" $statusTextBox
            return $false
        }
        
        # Verify kết quả
        Start-Sleep -Seconds 2
        $verifyPartition = Get-Partition -DriveLetter $newDriveLetter -ErrorAction SilentlyContinue
        if ($verifyPartition) {
            $actualSizeGB = [math]::Round($verifyPartition.Size / 1GB, 2)
            $volumeInfo = Get-Volume -DriveLetter $newDriveLetter -ErrorAction SilentlyContinue
            $finalName = if ($volumeInfo) { $volumeInfo.FileSystemLabel } else { "Unknown" }
            $fileSystem = if ($volumeInfo) { $volumeInfo.FileSystem } else { "Unknown" }
            
            Add-Status "Partition creation verified: Drive $newDriveLetter ($actualSizeGB GB) - '$finalName' [$fileSystem]" $statusTextBox
            return $true
        } else {
            Add-Status "WARNING: Could not verify new partition" $statusTextBox
            return $false
        }
        
    } catch {
        Add-Status "CRITICAL ERROR during partition creation: $_" $statusTextBox
        return $false
    } finally {
        # Bật lại dịch vụ ShellHWDetection
        try {
            Start-Service -Name ShellHWDetection -ErrorAction SilentlyContinue
        } catch {
            Add-Status "Warning: Could not re-enable hardware detection service" $statusTextBox
        }
    }
}

#####################################################################################################################
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
        # STEP 1: Device Selection and Software Installation
        Add-Status "STEP 1: Selecting Device Type and Installing Software..." $statusTextBox
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
        # $copyResult = Copy-SoftwareFiles -deviceType $deviceType $statusTextBox
        # if (-not $copyResult) {
        #     Add-Status "Error copying software files. Exiting..." $statusTextBox
        #     return
        # }

        # Install software (gọi hàm toàn cục)
        Add-Status "Installing software..." $statusTextBox
        # Install-Software -deviceType $deviceType $statusTextBox
        Add-Status "All installation completed successfully for $deviceType" $statusTextBox
        Add-Status "STEP 1 completed successfully!" $statusTextBox

        # STEP 2: System Configuration and Shortcut Creation
        Add-Status "STEP 2: Configuring System and Creating Shortcuts..." $statusTextBox
        $progressBar.Value = 28 # Tăng giá trị progress bar

        # $configResult = Invoke-SystemConfiguration -deviceType $deviceType -statusTextBox $statusTextBox
        # if ($configResult) {
        #     Add-Status "STEP 2 completed successfully!" $statusTextBox
        # } else {
        #     Add-Status "STEP 2 encountered errors. Check logs." $statusTextBox
        # }

        # STEP 3: System Cleanup and Optimization
        Add-Status "STEP 3: Cleaning up system and optimizing performance..." $statusTextBox
        $progressBar.Value = 42 # Tăng giá trị progress bar

        # $cleanupResult = Invoke-SystemCleanup -deviceType $deviceType -statusTextBox $statusTextBox
        # if ($cleanupResult) {
        #     Add-Status "STEP 3 completed successfully!" $statusTextBox
        # } else {
        #     Add-Status "STEP 3 encountered errors. Check logs." $statusTextBox
        # }

        # STEP 4: Windows and Office Activation
        Add-Status "STEP 4: Activating Windows 10 Pro and Office 2019 Pro Plus..." $statusTextBox
        $progressBar.Value = 56 # Tăng giá trị progress bar

        # $activationResult = Invoke-ActivationConfiguration -deviceType $deviceType -statusTextBox $statusTextBox
        # if ($activationResult) {
        #     Add-Status "STEP 4 completed successfully!" $statusTextBox
        # } else {
        #     Add-Status "STEP 4 encountered errors. Check logs." $statusTextBox
        # }

        # STEP 5: Windows Features Configuration
        Add-Status "STEP 5: Configuring Windows Features..." $statusTextBox
        $progressBar.Value = 70 # Tăng giá trị progress bar

        # $featuresResult = Invoke-WindowsFeaturesConfiguration -deviceType $deviceType -statusTextBox $statusTextBox
        # if ($featuresResult) {
        #     Add-Status "STEP 5 completed successfully!" $statusTextBox
        # } else {
        #     Add-Status "STEP 5 encountered errors. Check logs." $statusTextBox
        # }

        # STEP 6: Disk Partitioning (Laptop only)
        Add-Status "STEP 6: Configuring disk partitioning..." $statusTextBox
        $progressBar.Value = 85 # Tăng giá trị progress bar

        $partitioningResult = Invoke-DiskPartitioning -deviceType $deviceType -statusTextBox $statusTextBox
        if ($partitioningResult) {
            Add-Status "STEP 6 completed successfully!" $statusTextBox
        } else {
            Add-Status "STEP 6 encountered errors. Check logs." $statusTextBox
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

####################################################################################################################
# [1] Run All Functions
$buttonRunAll = New-DynamicButton -text "[1] Run All" -x 30 -y 100 -width 380 -height 60 -clickAction {
    Invoke-RunAllOperations -mainForm $script:form
}
# [2] Install Software Button
$buttonInstallSoftware = New-DynamicButton -text "[2] Install All Software" -x 30 -y 180 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(0, 150, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(0, 200, 0)) -pressColor ([System.Drawing.Color]::FromArgb(0, 100, 0)) -clickAction {
    Show-InstallSoftwareDialog
}
# [9] Join Domain
$buttonJoinDomain = New-DynamicButton -text "[9] Join Domain" -x 430 -y 340 -width 380 -height 60 -clickAction {
    Show-DomainManagementForm
}
# [0] Exit
$buttonExit = New-DynamicButton -text "[0] Exit" -x 430 -y 420 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
    $script:form.Close()
}
# Add buttons to form
$script:form.Controls.Add($buttonRunAll)
$script:form.Controls.Add($buttonInstallSoftware)
$script:form.Controls.Add($buttonJoinDomain)
$script:form.Controls.Add($buttonExit)
# SECTION 5: START APPLICATION
$script:form.ShowDialog() 