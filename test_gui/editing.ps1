# # ADMIN PRIVILEGES CHECK & INITIALIZATION
# if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
#     Write-Warning "This script requires administrative privileges. Attempting to restart with elevation..."

#     # Restart script with admin privileges
#     $scriptPath = $MyInvocation.MyCommand.Path
#     $arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""

#     Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs

#     # Exit the current non-elevated instance
#     exit
# }

# HIDING CONSOLE
try {
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    ' -ErrorAction SilentlyContinue
    
    $consolePtr = [Console.Window]::GetConsoleWindow()
    if ($consolePtr -ne [System.IntPtr]::Zero) {
        [Console.Window]::ShowWindow($consolePtr, 0)
    }
} catch {
}


# UTILITY FUNCTIONS
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

# Lệnh kiểm tra và tải thư viện Windows Forms
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
} catch {
    Write-Error "Failed to load Windows Forms: $_"
    Write-Host "Trying alternative method..."
    
    try {
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
        [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
        Write-Host "Windows Forms loaded via alternative method"
    } catch {
        Write-Error "Cannot load Windows Forms. Exiting..."
        Read-Host "Press Enter to exit"
        exit
    }
}

# Tạo form chính có thể thay đổi kích thước
$script:form = New-Object System.Windows.Forms.Form
$script:form.Text = "BAOPROVIP - SYSTEM MANAGEMENT"
$script:form.Size = New-Object System.Drawing.Size(850, 5)
$script:form.MinimumSize = New-Object System.Drawing.Size(800, 550)  # Kích thước tối thiểu
$script:form.StartPosition = "CenterScreen"
$script:form.BackColor = [System.Drawing.Color]::Black
$script:form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable  # CHỖ NÀY THAY ĐỔI
$script:form.MaximizeBox = $true  # Cho phép maximize

# Thêm màu gradient với resize handling
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

# Tiêu đề - RESPONSIVE
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "WELCOME TO BAROPROVIP"
$titleLabel.Font = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$titleLabel.Size = New-Object System.Drawing.Size($script:form.ClientSize.Width, 60)
$titleLabel.Location = New-Object System.Drawing.Point(0, 20)
$titleLabel.BackColor = [System.Drawing.Color]::Transparent
$titleLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
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

## STEP 0: WiFi AUTO-CONNECTION FUNCTION
function Invoke-WiFiAutoConnection {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Checking WiFi connection..." $statusTextBox
    
    try {
        # Phương pháp 1: Kiểm tra bằng InterfaceType (71 = Wireless80211)
        $wifiAdapter = Get-NetAdapter | Where-Object { $_.InterfaceType -eq 71 }
        
        if (-not $wifiAdapter) {
            # Phương pháp 2: Kiểm tra bằng PhysicalMediaType
            Add-Status "Method 1 failed, trying alternative detection..." $statusTextBox
            $wifiAdapter = Get-NetAdapter | Where-Object { 
                $_.PhysicalMediaType -eq 'Native 802.11' -or 
                $_.PhysicalMediaType -eq 'Wireless LAN' -or 
                $_.PhysicalMediaType -eq 'Wireless WAN' 
            }
        }
        
        if (-not $wifiAdapter) {
            # Phương pháp 3: Kiểm tra bằng InterfaceDescription
            Add-Status "Method 2 failed, trying description-based detection..." $statusTextBox
            $wifiAdapter = Get-NetAdapter | Where-Object { 
                $_.InterfaceDescription -like "*wireless*" -or 
                $_.InterfaceDescription -like "*wifi*" -or 
                $_.InterfaceDescription -like "*802.11*" -or
                $_.InterfaceDescription -like "*Wi-Fi*" -or
                $_.InterfaceDescription -like "*WLAN*"
            }
        }
        
        if (-not $wifiAdapter) {
            # Phương pháp 4: Kiểm tra bằng WMI
            Add-Status "Method 3 failed, trying WMI detection..." $statusTextBox
            try {
                $wmiWifiAdapters = Get-WmiObject -Class Win32_NetworkAdapter | Where-Object { 
                    $_.AdapterType -like "*802.11*" -or 
                    $_.Name -like "*wireless*" -or 
                    $_.Name -like "*wifi*" -or
                    $_.Name -like "*Wi-Fi*" -or
                    $_.Name -like "*WLAN*" -or
                    $_.Description -like "*wireless*"
                }
                
                if ($wmiWifiAdapters) {
                    Add-Status "WiFi adapter detected via WMI: $($wmiWifiAdapters[0].Name)" $statusTextBox
                    # Thử lấy lại bằng Get-NetAdapter với tên từ WMI
                    $wifiAdapter = Get-NetAdapter | Where-Object { $_.Name -eq $wmiWifiAdapters[0].NetConnectionID }
                }
            } catch {
                Add-Status "WMI detection failed: $_" $statusTextBox
            }
        }
        
        if (-not $wifiAdapter) {
            # Phương pháp 5: Kiểm tra service WLAN AutoConfig
            Add-Status "Method 4 failed, checking WLAN service..." $statusTextBox
            try {
                $wlanService = Get-Service -Name "WlanSvc" -ErrorAction SilentlyContinue
                if ($wlanService -and $wlanService.Status -eq "Running") {
                    Add-Status "WLAN service is running, but no adapter detected through PowerShell" $statusTextBox
                    Add-Status "Attempting direct netsh approach..." $statusTextBox
                    
                    # Thử sử dụng netsh để kiểm tra interfaces
                    $netshResult = netsh wlan show interfaces 2>$null
                    if ($netshResult -and $netshResult -notlike "*There is no wireless interface on the system*") {
                        Add-Status "WiFi interface detected via netsh, proceeding with connection..." $statusTextBox
                        # Tiếp tục với quá trình kết nối mà không cần PowerShell adapter object
                        $useNetshOnly = $true
                    } else {
                        Add-Status "No WiFi interface found via netsh either" $statusTextBox
                        Add-Status "No WiFi adapter found. Skipping WiFi connection..." $statusTextBox
                        return $true
                    }
                } else {
                    Add-Status "WLAN service not running. No WiFi capability detected." $statusTextBox
                    Add-Status "Skipping WiFi connection..." $statusTextBox
                    return $true
                }
            } catch {
                Add-Status "Service check failed: $_" $statusTextBox
                Add-Status "No WiFi adapter found. Skipping WiFi connection..." $statusTextBox
                return $true
            }
        }
        
        if ($wifiAdapter -and -not $useNetshOnly) {
            Add-Status "WiFi adapter found: $($wifiAdapter.Name) - $($wifiAdapter.InterfaceDescription)" $statusTextBox
        }
        
        # Kiểm tra xem đã kết nối WiFi "VietUnion_5.0GHz" chưa
        try {
            $currentConnection = netsh wlan show interfaces | Select-String "SSID" | Select-String "VietUnion_5.0GHz"
            
            if ($currentConnection) {
                Add-Status "Already connected to 'VietUnion_5.0GHz' WiFi. Skipping..." $statusTextBox
                return $true
            }
        } catch {
            Add-Status "Could not check current connection, proceeding with connection attempt..." $statusTextBox
        }
        
        # Thông tin WiFi
        $SSID = "VietUnion_5.0GHz"
        $Password = "Pay00@17Years$"
        $profileFile = "$env:TEMP\VietUnion_5.0GHz_profile.xml"
        
        # Tạo hex cho SSID
        $SSIDHEX = ($SSID.ToCharArray() | ForEach-Object {'{0:X}' -f ([int]$_)}) -join ''
        
        # Tạo XML profile cho WiFi
        $xmlContent = @"
<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>$SSID</name>
    <SSIDConfig>
        <SSID>
            <hex>$SSIDHEX</hex>
            <name>$SSID</name>
        </SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>WPA2PSK</authentication>
                <encryption>AES</encryption>
                <useOneX>false</useOneX>
            </authEncryption>
            <sharedKey>
                <keyType>passPhrase</keyType>
                <protected>false</protected>
                <keyMaterial>$Password</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>
"@
        
        # Ghi XML profile ra file
        try {
            $xmlContent | Out-File -FilePath $profileFile -Encoding UTF8
        } catch {
            Add-Status "ERROR: Could not create WiFi profile file: $_" $statusTextBox
            return $false
        }
        
        # Thêm profile WiFi
        try {
            $addResult = Start-Process -FilePath "netsh" -ArgumentList "wlan add profile filename=`"$profileFile`"" -Wait -PassThru -WindowStyle Hidden
            
            if ($addResult.ExitCode -eq 0) {
            } else {
                Add-Status "Warning: WiFi profile add returned exit code $($addResult.ExitCode)" $statusTextBox
            }
        } catch {
            Add-Status "ERROR adding WiFi profile: $_" $statusTextBox
        }
        
        # Kết nối WiFi
        Add-Status "Connecting to 'VietUnion_5.0GHz' WiFi..." $statusTextBox
        try {
            $connectResult = Start-Process -FilePath "netsh" -ArgumentList "wlan connect name=`"$SSID`"" -Wait -PassThru -WindowStyle Hidden
            
            if ($connectResult.ExitCode -eq 0) {          
                # Đợi một chút để kết nối ổn định
                Add-Status "Waiting for connection to establish..." $statusTextBox
                Start-Sleep -Seconds 5
                
                # Xác minh kết nối
                try {
                    $verifyConnection = netsh wlan show interfaces | Select-String "SSID" | Select-String "VietUnion_5.0GHz"
                    if ($verifyConnection) {
                    } else {
                        Add-Status "Warning: Could not verify WiFi connection to 'VietUnion_5.0GHz'" $statusTextBox
                        # Kiểm tra xem có kết nối WiFi nào không
                        $anyConnection = netsh wlan show interfaces | Select-String "State" | Select-String "connected"
                        if ($anyConnection) {
                            Add-Status "Device is connected to a different WiFi network" $statusTextBox
                        } else {
                            Add-Status "Device is not connected to any WiFi network" $statusTextBox
                        }
                    }
                } catch {
                    Add-Status "Could not verify connection status" $statusTextBox
                }
            } else {
                Add-Status "Warning: WiFi connection command returned exit code $($connectResult.ExitCode)" $statusTextBox
            }
        } catch {
            Add-Status "ERROR connecting to WiFi: $_" $statusTextBox
        }
        
        # Xóa file profile tạm
        try {
            if (Test-Path $profileFile) {
                Remove-Item -Path $profileFile -Force -ErrorAction SilentlyContinue
            }
        } catch {
            # Ignore cleanup errors
        }
        
        return $true
        
    } catch {
        Add-Status "ERROR during WiFi connection: $_" $statusTextBox
        return $false
    }
}

## STEP 1: Software Installation & Rename Functions
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
        
        # 1. Check and uninstall OneDrive if present - SIMPLIFIED STATUS VERSION
$oneDrivePaths = @(
    "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe",
    "$env:PROGRAMFILES\Microsoft OneDrive\OneDrive.exe",
    "$env:PROGRAMFILES(x86)\Microsoft OneDrive\OneDrive.exe"
)

$oneDriveFound = $false
$oneDriveExecutable = $null

# Tìm OneDrive executable
foreach ($path in $oneDrivePaths) {
    if (Test-Path $path) {
        $oneDriveFound = $true
        $oneDriveExecutable = $path
        break
    }
}

if ($oneDriveFound) {
    Add-Status "OneDrive found. Uninstalling..." $statusTextBox
    
    try {
        # Force kill OneDrive processes
        $oneDriveProcesses = @("OneDrive", "OneDriveSetup", "FileCoAuth", "OneDriveStandaloneUpdater", "OneDriveUpdaterService")
        
        foreach ($processName in $oneDriveProcesses) {
            try {
                $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
                if ($processes) {
                    foreach ($proc in $processes) {
                        $proc.Kill()
                        $proc.WaitForExit(5000)
                    }
                }
            } catch {
                # Silent error handling
            }
        }
        
        # Stop OneDrive services
        $services = @("OneDrive Updater Service")
        foreach ($serviceName in $services) {
            try {
                $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
                if ($service -and $service.Status -eq 'Running') {
                    Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
                }
            } catch {
                # Silent error handling
            }
        }
        
        Start-Sleep -Seconds 2
        
        # Try uninstall with OneDriveSetup.exe
        $uninstallSuccess = $false
        $setupPaths = @(
            "$env:SYSTEMROOT\SysWOW64\OneDriveSetup.exe",
            "$env:SYSTEMROOT\System32\OneDriveSetup.exe"
        )
        
        foreach ($setupPath in $setupPaths) {
            if (Test-Path $setupPath) {
                try {
                    $result = Start-Process -FilePath $setupPath -ArgumentList "/uninstall /allusers" -Wait -PassThru -WindowStyle Hidden
                    
                    if ($result.ExitCode -eq 0) {
                        $uninstallSuccess = $true
                        break
                    }
                } catch {
                    # Try next method
                    continue
                }
            }
        }
        
        # Manual cleanup if uninstall failed
        if (-not $uninstallSuccess) {
            # Clean registry entries
            $registryPaths = @(
                "HKCU:\Software\Microsoft\OneDrive",
                "HKLM:\SOFTWARE\Microsoft\OneDrive"
            )
            
            foreach ($regPath in $registryPaths) {
                if (Test-Path $regPath) {
                    Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
            
            # Remove from startup
            try {
                Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "OneDrive" -ErrorAction SilentlyContinue
            } catch {
                # Silent error handling
            }
            
            # Clean folders
            $oneDriveFolders = @(
                "$env:LOCALAPPDATA\Microsoft\OneDrive",
                "$env:PROGRAMDATA\Microsoft OneDrive",
                "$env:USERPROFILE\OneDrive",
                "$env:PROGRAMFILES\Microsoft OneDrive",
                "$env:PROGRAMFILES(x86)\Microsoft OneDrive"
            )
            
            foreach ($folder in $oneDriveFolders) {
                if (Test-Path $folder) {
                    try {
                        takeown /f "$folder" /r /d y 2>$null | Out-Null
                        icacls "$folder" /grant administrators:F /t 2>$null | Out-Null
                        
                        Get-ChildItem -Path $folder -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                            try {
                                $_.Attributes = 'Normal'
                            } catch {
                                # Silent error handling
                            }
                        }
                        
                        Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
                    } catch {
                        # Silent error handling
                    }
                }
            }
            
            # Remove from File Explorer navigation pane
            try {
                $regPath1 = "HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
                if (Test-Path $regPath1) {
                    Set-ItemProperty -Path $regPath1 -Name "System.IsPinnedToNameSpaceTree" -Value 0 -Type DWord -ErrorAction SilentlyContinue
                }
                
                $regPath2 = "HKCR:\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
                if (Test-Path $regPath2) {
                    Set-ItemProperty -Path $regPath2 -Name "System.IsPinnedToNameSpaceTree" -Value 0 -Type DWord -ErrorAction SilentlyContinue
                }
            } catch {
                # Silent error handling
            }
        }
        
        Add-Status "OneDrive uninstalled successfully!" $statusTextBox
        
    } catch {
        Add-Status "OneDrive uninstalled successfully!" $statusTextBox
    }
} else {
    Add-Status "OneDrive:     Has Not installed. Skipping..." $statusTextBox
}

        
        # 2. Install 7-Zip - FIXED VERSION
        $sevenZipPaths = @(
            "C:\Program Files\7-Zip\7z.exe",
            "C:\Program Files (x86)\7-Zip\7z.exe"
        )

        $sevenZipInstalled = $false
        foreach ($path in $sevenZipPaths) {
            if (Test-Path $path) {
                $sevenZipInstalled = $true
                break
            }
        }

        if (-not $sevenZipInstalled) {
            # Tìm file installer với nhiều pattern
            $sevenZipFiles = @()
            $searchPatterns = @("7z*.exe", "7-Zip*.exe", "7zip*.exe")
            
            foreach ($pattern in $searchPatterns) {
                $foundFiles = Get-ChildItem -Path $setupDir -Name $pattern -ErrorAction SilentlyContinue
                if ($foundFiles) {
                    $sevenZipFiles += $foundFiles
                    break
                }
            }
            
            if ($sevenZipFiles.Count -gt 0) {
                $sevenZipInstaller = "$setupDir\$($sevenZipFiles[0])"
                Add-Status "Installing 7-Zip..." $statusTextBox
                
                try {
                    # Cài đặt với kiểm tra exit code
                    $result = Start-Process -FilePath $sevenZipInstaller -ArgumentList "/S" -Wait -PassThru -WindowStyle Hidden
                
                    if ($result.ExitCode -eq 0) {
                        Add-Status "7-Zip installed successfully!" $statusTextBox
                    } else {
                        Add-Status "WARNING: 7-Zip EXE installation returned exit code: $($result.ExitCode)" $statusTextBox
                    }
                } catch {
                    Add-Status "ERROR: 7-Zip installation failed: $_" $statusTextBox
                }
            }
        } else {
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
            $lapsInstaller = "$setupDir\LAPS_x64.msi"
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
        
        # 5. Install Foxit Reader - COMPLETELY SILENT VERSION
        $foxitCheck = @(
            "C:\Program Files (x86)\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe",
            "C:\Program Files\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe",
            "C:\Program Files (x86)\Foxit Software\Foxit Reader\FoxitReader.exe",
            "C:\Program Files\Foxit Software\Foxit Reader\FoxitReader.exe"
        )

        $foxitInstalled = $false
        foreach ($path in $foxitCheck) {
            if (Test-Path $path) {
                $foxitInstalled = $true
                break
            }
        }

        if (-not $foxitInstalled) {
            # Tìm file installer với nhiều pattern
            $foxitFiles = @()
            $searchPatterns = @("FoxitPDFReader*.exe", "FoxitReader*.exe", "Foxit*.exe")
            
            foreach ($pattern in $searchPatterns) {
                $foundFiles = Get-ChildItem -Path $setupDir -Name $pattern -ErrorAction SilentlyContinue
                if ($foundFiles) {
                    $foxitFiles += $foundFiles
                    break
                }
            }
            
            if ($foxitFiles.Count -gt 0) {
                $foxitPath = "$setupDir\$($foxitFiles[0])"
                Add-Status "Installing Foxit Reader..." $statusTextBox
                
                try {
                    # SỬ DỤNG /VERYSILENT ĐỂ ẨN HOÀN TOÀN BẢNG SETUP
                    $result = Start-Process -FilePath $foxitPath -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART" -Wait -PassThru -WindowStyle Hidden
                    
                    if ($result.ExitCode -eq 0) {
                        Add-Status "Foxit Reader installed successfully!" $statusTextBox
                    } else {
                        Add-Status "ERROR: Foxit Reader installation failed (Exit code: $($result.ExitCode))" $statusTextBox
                    } 
                } catch {
                    Add-Status "ERROR: Foxit Reader installation failed: $_" $statusTextBox
                }
            } else {
                Add-Status "ERROR: Foxit Reader installer not found in $setupDir" $statusTextBox
            }
        } else {
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

# Function to show Rename Device dialog
function Show-RenameDeviceDialog {
    # Hide the main menu
    Hide-MainMenu

    # Create rename device form
    $renameForm = New-Object System.Windows.Forms.Form
    $renameForm.Text = "Rename Device"
    $renameForm.Size = New-Object System.Drawing.Size(500, 420)  # Tăng chiều cao
    $renameForm.StartPosition = "CenterScreen"
    $renameForm.BackColor = [System.Drawing.Color]::Black
    $renameForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $renameForm.MaximizeBox = $false
    $renameForm.MinimizeBox = $false

    # Add gradient background
    $renameForm.Add_Paint({
        $graphics = $_.Graphics
        $rect = New-Object System.Drawing.Rectangle(0, 0, $renameForm.Width, $renameForm.Height)
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
    $titleLabel.Text = "RENAME DEVICE"
    $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(500, 35)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 18, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $renameForm.Controls.Add($titleLabel)

    # Current computer name
    $currentName = $env:COMPUTERNAME
    $currentNameLabel = New-Object System.Windows.Forms.Label
    $currentNameLabel.Text = "Current Computer Name: $currentName"
    $currentNameLabel.Location = New-Object System.Drawing.Point(20, 70)
    $currentNameLabel.Size = New-Object System.Drawing.Size(460, 30)
    $currentNameLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $currentNameLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $currentNameLabel.BackColor = [System.Drawing.Color]::Transparent
    $renameForm.Controls.Add($currentNameLabel)

    # GROUPBOX CHO NAMING OPTIONS
    $namingGroupBox = New-Object System.Windows.Forms.GroupBox
    $namingGroupBox.Text = "Select Naming Option"
    $namingGroupBox.Location = New-Object System.Drawing.Point(20, 110)
    $namingGroupBox.Size = New-Object System.Drawing.Size(460, 80)
    $namingGroupBox.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 0)  # Vàng sáng
    $namingGroupBox.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $namingGroupBox.BackColor = [System.Drawing.Color]::Transparent
    $renameForm.Controls.Add($namingGroupBox)

    # Radio buttons TRONG GROUPBOX
    $radioDesktop = New-Object System.Windows.Forms.RadioButton
    $radioDesktop.Text = "Desktop"
    $radioDesktop.Location = New-Object System.Drawing.Point(15, 25)  # Relative to GroupBox
    $radioDesktop.Size = New-Object System.Drawing.Size(140, 30)
    $radioDesktop.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $radioDesktop.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $radioDesktop.BackColor = [System.Drawing.Color]::Transparent
    $radioDesktop.Checked = $true
    $namingGroupBox.Controls.Add($radioDesktop)  # Add to GroupBox

    $radioLaptop = New-Object System.Windows.Forms.RadioButton
    $radioLaptop.Text = "Laptop"
    $radioLaptop.Location = New-Object System.Drawing.Point(160, 25)  # Relative to GroupBox
    $radioLaptop.Size = New-Object System.Drawing.Size(130, 30)
    $radioLaptop.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $radioLaptop.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $radioLaptop.BackColor = [System.Drawing.Color]::Transparent
    $namingGroupBox.Controls.Add($radioLaptop)  # Add to GroupBox

    # CUSTOM NAME RADIO BUTTON TRONG GROUPBOX
    $radioCustom = New-Object System.Windows.Forms.RadioButton
    $radioCustom.Text = "Custom"
    $radioCustom.Location = New-Object System.Drawing.Point(295, 25)  # Relative to GroupBox
    $radioCustom.Size = New-Object System.Drawing.Size(150, 30)
    $radioCustom.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 255)
    $radioCustom.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $radioCustom.BackColor = [System.Drawing.Color]::Transparent
    $namingGroupBox.Controls.Add($radioCustom)  # Add to GroupBox

    # New name input - ĐIỀU CHỈNH VỊ TRÍ
    $newNameLabel = New-Object System.Windows.Forms.Label
    $newNameLabel.Text = "Enter new device name:"
    $newNameLabel.Location = New-Object System.Drawing.Point(20, 205)  # Điều chỉnh vị trí
    $newNameLabel.Size = New-Object System.Drawing.Size(460, 30)
    $newNameLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 0)
    $newNameLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $newNameLabel.BackColor = [System.Drawing.Color]::Transparent
    $renameForm.Controls.Add($newNameLabel)

    # TEXTBOX CẢI THIỆN
    $newNameTextBox = New-Object System.Windows.Forms.TextBox
    $newNameTextBox.Location = New-Object System.Drawing.Point(20, 240)  # Điều chỉnh vị trí
    $newNameTextBox.Size = New-Object System.Drawing.Size(350, 30)
    $newNameTextBox.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $newNameTextBox.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $newNameTextBox.ForeColor = [System.Drawing.Color]::White
    $newNameTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $renameForm.Controls.Add($newNameTextBox)

    # Preview label - ĐIỀU CHỈNH VỊ TRÍ
    $previewLabel = New-Object System.Windows.Forms.Label
    $previewLabel.Text = "New name will be: HOD"
    $previewLabel.Location = New-Object System.Drawing.Point(20, 280)  # Điều chỉnh vị trí
    $previewLabel.Size = New-Object System.Drawing.Size(460, 30)
    $previewLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 128)
    $previewLabel.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $previewLabel.BackColor = [System.Drawing.Color]::Transparent
    $renameForm.Controls.Add($previewLabel)

    # CẬP NHẬT HÀM UPDATE-PREVIEW
    function Update-Preview {
        $inputText = $newNameTextBox.Text.Trim()
        
        if ($radioCustom.Checked) {
            # Custom name mode
            $previewText = if ($inputText) { $inputText } else { "[Enter custom name]" }
            $previewLabel.Text = "New name will be: $previewText"
            $previewLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 255)
        } else {
            # Desktop/Laptop mode
            $prefix = if ($radioDesktop.Checked) { "HOD" } else { "HOL" }
            $previewText = if ($inputText) { "$prefix$inputText" } else { $prefix }
            $previewLabel.Text = "New name will be: $previewText"
            $previewLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 128)
        }
    }

    $newNameTextBox.Add_TextChanged({ Update-Preview })
    $radioDesktop.Add_CheckedChanged({ Update-Preview })
    $radioLaptop.Add_CheckedChanged({ Update-Preview })
    $radioCustom.Add_CheckedChanged({ Update-Preview })

    # BUTTONS - ĐIỀU CHỈNH VỊ TRÍ
    $renameButton = New-Object System.Windows.Forms.Button
    $renameButton.Text = "Rename Device"
    $renameButton.Location = New-Object System.Drawing.Point(150, 320)  # Điều chỉnh vị trí
    $renameButton.Size = New-Object System.Drawing.Size(130, 40)
    $renameButton.BackColor = [System.Drawing.Color]::FromArgb(0, 180, 0)
    $renameButton.ForeColor = [System.Drawing.Color]::White
    $renameButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $renameButton.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $renameButton.FlatAppearance.BorderSize = 1
    $renameButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
    $renameButton.Add_Click({
        $inputName = $newNameTextBox.Text.Trim()
        if ([string]::IsNullOrEmpty($inputName)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please enter a device name.",
                "Missing Input",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        # XỬ LÝ TÊN MỚI DỰA TRÊN OPTION
        if ($radioCustom.Checked) {
            $newComputerName = $inputName
            
            # Validate custom name
            if ($inputName.Length -gt 15) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Computer name cannot be longer than 15 characters.",
                    "Invalid Name",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }
            
            if ($inputName -match '[^a-zA-Z0-9\-]') {
                [System.Windows.Forms.MessageBox]::Show(
                    "Computer name can only contain letters, numbers, and hyphens.",
                    "Invalid Characters",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }
        } else {
            $prefix = if ($radioDesktop.Checked) { "HOD" } else { "HOL" }
            $newComputerName = "$prefix$inputName"
        }

        if ($newComputerName -eq $currentName) {
            [System.Windows.Forms.MessageBox]::Show(
                "New name is the same as current name.",
                "No Change",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }

        # Confirm rename
        $deviceTypeText = if ($radioCustom.Checked) { "Custom Name" } elseif ($radioDesktop.Checked) { "Desktop" } else { "Laptop" }
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to rename the computer to '$newComputerName'?`n`nDevice Type: $deviceTypeText`nThe computer will restart after renaming.",
            "Confirm Rename",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Rename-Computer -NewName $newComputerName -Force -Restart
                [System.Windows.Forms.MessageBox]::Show(
                    "Computer will be renamed to '$newComputerName' and restarted.",
                    "Rename Successful",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                $renameForm.Close()
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to rename computer: $_",
                    "Rename Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
    })
    $renameForm.Controls.Add($renameButton)

    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(290, 320)  # Điều chỉnh vị trí
    $cancelButton.Size = New-Object System.Drawing.Size(100, 40)
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $cancelButton.ForeColor = [System.Drawing.Color]::White
    $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelButton.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $cancelButton.FlatAppearance.BorderSize = 1
    $cancelButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(255, 0, 0)
    $cancelButton.Add_Click({
        $renameForm.Close()
    })
    $renameForm.Controls.Add($cancelButton)

    # Close button (X)
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "✕"
    $closeButton.Location = New-Object System.Drawing.Point(10, 10)
    $closeButton.Size = New-Object System.Drawing.Size(30, 30)
    $closeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $closeButton.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $closeButton.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $closeButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $closeButton.FlatAppearance.BorderSize = 1
    $closeButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $closeButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $closeButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $closeButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $closeButton.Add_Click({
        $renameForm.Close()
    })
    $renameForm.Controls.Add($closeButton)

    # Keyboard events
    $renameForm.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $renameForm.Close()
        } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $renameButton.PerformClick()
        }
    })

    $renameForm.KeyPreview = $true

    # Focus on text box when form shows
    $renameForm.Add_Shown({
        $newNameTextBox.Focus()
    })

    # When form closes, show main menu
    $renameForm.Add_FormClosed({
        Show-MainMenu
    })

    # Show the dialog
    $renameForm.ShowDialog()
}

## STEP 2: Hostname Configuration Functions
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

## STEP 3: POWER OPTIONS FUNCTIONS
function Invoke-SystemCleanup {
    param (
        [string]$deviceType,
        [System.Windows.Forms.TextBox]$statusTextBox
    )
    
    try {
        Add-Status "Starting system cleanup and optimization..." $statusTextBox
        
        # --- 1. System File Cleanup ---
        Invoke-FileCleanup $statusTextBox
       
        # --- 2. Taskbar Customization ---
        Invoke-TaskbarCustomization $statusTextBox 

        # --- 4. Startup Program Management ---
        Invoke-StartupOptimization $statusTextBox
        
        # --- 5.
        Invoke-DiskOptimization $statusTextBox

        # --- 6. Timezone Configuration ---
        Invoke-TimezoneConfiguration $statusTextBox
        
        # --- 7. Power Options Configuration ---
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

# Hàm tối ưu hóa startup programs
function Invoke-StartupOptimization {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Disabling unnecessary startup programs..." $statusTextBox
    
    $startupPrograms = @("Skype for Desktop", "Microsoft Co-Pilot", "Microsoft Edge")
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

# Hàm tối ưu hóa disk
function Invoke-DiskOptimization {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    try {        
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

# Hàm cấu hình múi giờ
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

# Hàm cấu hình power options
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
        }
        
        # Tắt timeout cho màn hình và sleep
        powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 0 | Out-Null
        powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 0 | Out-Null
        powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_SLEEP STANDBYIDLE 0 | Out-Null
        powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_SLEEP STANDBYIDLE 0 | Out-Null
        
        # Áp dụng thay đổi
        powercfg /SETACTIVE SCHEME_CURRENT | Out-Null
        Add-Status "Power options configured to  'Do Nothing' completed successfully!" $statusTextBox
    } catch {
        Add-Status "Warning: Could not configure power options: $_" $statusTextBox
    }
}

# Function to customize taskbar - Windows 10 & 11 compatible
function Invoke-TaskbarCustomization {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Customizing taskbar settings..." $statusTextBox
    
    try {
        # Detect Windows version
        $osVersion = (Get-CimInstance Win32_OperatingSystem).Caption
        $isWindows11 = $osVersion -like "*Windows 11*"
        
        if ($isWindows11) {
            Add-Status "Detected Windows 11 - applying specific customizations..." $statusTextBox
        } else {
            Add-Status "Detected Windows 10 - applying specific customizations..." $statusTextBox
        }
        
        # 1. UNPIN MICROSOFT STORE
        Add-Status "Unpinning Microsoft Store from taskbar..." $statusTextBox
        try {
            # Method 1: PowerShell App Package removal from taskbar
            $storeAppId = "Microsoft.WindowsStore_8wekyb3d8bbwe!App"
            
            # Windows 11 method
            if ($isWindows11) {
                # Remove from taskbar via registry
                $taskbarRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
                if (Test-Path $taskbarRegPath) {
                    Remove-ItemProperty -Path $taskbarRegPath -Name "Favorites" -ErrorAction SilentlyContinue
                }
                
                # Use PowerShell to unpin
                $shell = New-Object -ComObject Shell.Application
                $folder = $shell.Namespace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}')
                if ($folder) {
                    $item = $folder.ParseName("Microsoft Store")
                    if ($item) {
                        $item.InvokeVerb("Unpin from tas&kbar")
                        Add-Status "Microsoft Store unpinned from taskbar" $statusTextBox
                    }
                }
            } else {
                # Windows 10 method
                $shell = New-Object -ComObject Shell.Application
                $folder = $shell.Namespace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}')
                if ($folder) {
                    $item = $folder.ParseName("Microsoft Store")
                    if ($item) {
                        $item.InvokeVerb("Unpin from tas&kbar")
                        Add-Status "Microsoft Store unpinned from taskbar" $statusTextBox
                    }
                }
            }
        } catch {
            Add-Status "Could not unpin Microsoft Store: $_" $statusTextBox
        }
        
        # 2. DISABLE/UNPIN MS COPILOT (Windows 11 specific)
        if ($isWindows11) {
            Add-Status "Disabling MS Copilot..." $statusTextBox
            try {
                # Disable Copilot via registry
                $copilotRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
                if (-not (Test-Path $copilotRegPath)) {
                    New-Item -Path $copilotRegPath -Force | Out-Null
                }
                
                # Disable Copilot button on taskbar
                Set-ItemProperty -Path $copilotRegPath -Name "ShowCopilotButton" -Value 0 -Type DWord -ErrorAction SilentlyContinue
                
                # Disable Copilot via Group Policy equivalent
                $copilotPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot"
                if (-not (Test-Path $copilotPolicyPath)) {
                    New-Item -Path $copilotPolicyPath -Force -ErrorAction SilentlyContinue | Out-Null
                }
                Set-ItemProperty -Path $copilotPolicyPath -Name "TurnOffWindowsCopilot" -Value 1 -Type DWord -ErrorAction SilentlyContinue
                
                Add-Status "MS Copilot disabled successfully" $statusTextBox
            } catch {
                Add-Status "Could not disable MS Copilot: $_" $statusTextBox
            }
        } else {
            Add-Status "MS Copilot not applicable for Windows 10" $statusTextBox
        }
        
        # 3. DISABLE WIDGETS (Windows 11) / NEWS AND INTERESTS (Windows 10)
        Add-Status "Disabling Widgets/News and Interests..." $statusTextBox
        try {
            $explorerRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            
            if ($isWindows11) {
                # Windows 11 - Disable Widgets
                Set-ItemProperty -Path $explorerRegPath -Name "TaskbarDa" -Value 0 -Type DWord -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $explorerRegPath -Name "TaskbarWidgets" -Value 0 -Type DWord -ErrorAction SilentlyContinue
                Add-Status "Widgets disabled for Windows 11" $statusTextBox
            } else {
                # Windows 10 - Disable News and Interests
                Set-ItemProperty -Path $explorerRegPath -Name "TaskbarDa" -Value 0 -Type DWord -ErrorAction SilentlyContinue
                
                # Additional registry for News and Interests
                $feedsRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Feeds"
                if (-not (Test-Path $feedsRegPath)) {
                    New-Item -Path $feedsRegPath -Force | Out-Null
                }
                Set-ItemProperty -Path $feedsRegPath -Name "ShellFeedsTaskbarViewMode" -Value 2 -Type DWord -ErrorAction SilentlyContinue
                
                Add-Status "News and Interests disabled for Windows 10" $statusTextBox
            }
        } catch {
            Add-Status "Could not disable Widgets/News and Interests: $_" $statusTextBox
        }
        
        # 4. HIDE TASK VIEW BUTTON
        Add-Status "Hiding Task View button..." $statusTextBox
        try {
            $explorerRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            
            # Hide Task View button (works for both Windows 10 and 11)
            Set-ItemProperty -Path $explorerRegPath -Name "ShowTaskViewButton" -Value 0 -Type DWord -ErrorAction SilentlyContinue
            
            Add-Status "Task View button hidden successfully" $statusTextBox
        } catch {
            Add-Status "Could not hide Task View button: $_" $statusTextBox
        }
        
        # 5. ADDITIONAL TASKBAR CUSTOMIZATIONS
        Add-Status "Applying additional taskbar customizations..." $statusTextBox
        try {
            $explorerRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            
            # Hide People button (Windows 10)
            if (-not $isWindows11) {
                Set-ItemProperty -Path $explorerRegPath -Name "PeopleBand" -Value 0 -Type DWord -ErrorAction SilentlyContinue
                Add-Status "People button hidden (Windows 10)" $statusTextBox
            }
            
            # Hide Meet Now button (Windows 10)
            if (-not $isWindows11) {
                $meetNowRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
                if (-not (Test-Path $meetNowRegPath)) {
                    New-Item -Path $meetNowRegPath -Force | Out-Null
                }
                Set-ItemProperty -Path $meetNowRegPath -Name "HideSCAMeetNow" -Value 1 -Type DWord -ErrorAction SilentlyContinue
                Add-Status "Meet Now button hidden (Windows 10)" $statusTextBox
            }
            
            # Windows 11 specific - Hide Chat button
            if ($isWindows11) {
                Set-ItemProperty -Path $explorerRegPath -Name "TaskbarMn" -Value 0 -Type DWord -ErrorAction SilentlyContinue
                Add-Status "Chat button hidden (Windows 11)" $statusTextBox
            }
            
        } catch {
            Add-Status "Could not apply additional customizations: $_" $statusTextBox
        }
        
        # 6. RESTART EXPLORER TO APPLY CHANGES
        Add-Status "Restarting Windows Explorer to apply changes..." $statusTextBox
        try {
            # Kill explorer process
            Stop-Process -Name "explorer" -Force -ErrorAction SilentlyContinue
            
            # Wait a moment
            Start-Sleep -Seconds 2
            
            # Start explorer again
            Start-Process "explorer.exe"
            
            Add-Status "Windows Explorer restarted successfully" $statusTextBox
        } catch {
            Add-Status "Could not restart Explorer: $_" $statusTextBox
        }
        
        Add-Status "Taskbar customization completed successfully!" $statusTextBox
        
    } catch {
        Add-Status "ERROR during taskbar customization: $_" $statusTextBox
    }
}

# Alternative method using PowerShell 7+ and Windows Terminal commands
function Invoke-AdvancedTaskbarCustomization {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Applying advanced taskbar customizations..." $statusTextBox
    
    try {
        # Unpin Microsoft Store using PowerShell
        Add-Status "Unpinning Microsoft Store using PowerShell method..." $statusTextBox
        
        $unpinScript = @'
$shell = New-Object -ComObject Shell.Application
$folder = $shell.Namespace("shell:::{4234d49b-0245-4df3-b780-3893943456e1}")
$items = $folder.Items()
foreach ($item in $items) {
    if ($item.Name -eq "Microsoft Store") {
        $verbs = $item.Verbs()
        foreach ($verb in $verbs) {
            if ($verb.Name -like "*Unpin*taskbar*" -or $verb.Name -like "*Unpin*tas&kbar*") {
                $verb.DoIt()
                break
            }
        }
        break
    }
}
'@
        
        try {
            Invoke-Expression $unpinScript
            Add-Status "Microsoft Store unpinned via PowerShell method" $statusTextBox
        } catch {
            Add-Status "PowerShell unpin method failed: $_" $statusTextBox
        }
        
        # Force registry changes
        Add-Status "Forcing registry changes..." $statusTextBox
        
        # Create comprehensive registry script
        $regScript = @"
Windows Registry Editor Version 5.00

; Hide Task View Button
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
"ShowTaskViewButton"=dword:00000000

; Hide Widgets (Windows 11) / News and Interests (Windows 10)
"TaskbarDa"=dword:00000000
"TaskbarWidgets"=dword:00000000

; Hide Copilot Button (Windows 11)
"ShowCopilotButton"=dword:00000000

; Hide Chat Button (Windows 11)
"TaskbarMn"=dword:00000000

; Hide People Button (Windows 10)
"PeopleBand"=dword:00000000

; Disable News and Interests (Windows 10)
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Feeds]
"ShellFeedsTaskbarViewMode"=dword:00000002

; Hide Meet Now (Windows 10)
[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer]
"HideSCAMeetNow"=dword:00000001

; Disable Copilot Policy (Windows 11)
[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot]
"TurnOffWindowsCopilot"=dword:00000001
"@
        
        # Write and apply registry file
        $regFile = "$env:TEMP\taskbar_customization.reg"
        $regScript | Out-File -FilePath $regFile -Encoding ASCII
        
        try {
            Start-Process -FilePath "reg" -ArgumentList "import `"$regFile`"" -Wait -WindowStyle Hidden
            Add-Status "Registry changes applied successfully" $statusTextBox
            
            # Clean up
            Remove-Item -Path $regFile -Force -ErrorAction SilentlyContinue
        } catch {
            Add-Status "Could not apply registry changes: $_" $statusTextBox
        }
        
    } catch {
        Add-Status "ERROR in advanced taskbar customization: $_" $statusTextBox
    }
}

## STEP 4: Activation Configuration Functions
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

# Hàm 
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

# 
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

## STEP 5: Windows Features Configuration Functions
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
function Show-TurnOnFeaturesDialog {
    # Hide the main menu
    Hide-MainMenu

    # Create features configuration form
    $featuresForm = New-Object System.Windows.Forms.Form
    $featuresForm.Text = "Windows Features Configuration"
    $featuresForm.Size = New-Object System.Drawing.Size(600, 500)
    $featuresForm.StartPosition = "CenterScreen"
    $featuresForm.BackColor = [System.Drawing.Color]::Black
    $featuresForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $featuresForm.MaximizeBox = $false
    $featuresForm.MinimizeBox = $false

    # Add gradient background
    $featuresForm.Add_Paint({
        $graphics = $_.Graphics
        $rect = New-Object System.Drawing.Rectangle(0, 0, $featuresForm.Width, $featuresForm.Height)
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
    $titleLabel.Text = "WINDOWS FEATURES CONFIGURATION"
    $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(600, 35)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $titleLabel.BackColor = [System.Drawing.Color]::Transparent
    $featuresForm.Controls.Add($titleLabel)

    # Status textbox
    $featuresStatusTextBox = New-Object System.Windows.Forms.TextBox
    $featuresStatusTextBox.Multiline = $true
    $featuresStatusTextBox.ScrollBars = "Vertical"
    $featuresStatusTextBox.Location = New-Object System.Drawing.Point(20, 70)
    $featuresStatusTextBox.Size = New-Object System.Drawing.Size(560, 300)
    $featuresStatusTextBox.BackColor = [System.Drawing.Color]::Black
    $featuresStatusTextBox.ForeColor = [System.Drawing.Color]::Lime
    $featuresStatusTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $featuresStatusTextBox.ReadOnly = $true
    $featuresStatusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $featuresStatusTextBox.Text = "Ready to configure Windows Features..."
    $featuresForm.Controls.Add($featuresStatusTextBox)

    # Start Configuration button
    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Text = "Start Configuration"
    $startButton.Location = New-Object System.Drawing.Point(150, 390)
    $startButton.Size = New-Object System.Drawing.Size(150, 40)
    $startButton.BackColor = [System.Drawing.Color]::FromArgb(0, 180, 0)
    $startButton.ForeColor = [System.Drawing.Color]::White
    $startButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $startButton.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $startButton.FlatAppearance.BorderSize = 1
    $startButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
    $startButton.Add_Click({
        try {
            $startButton.Enabled = $false
            $startButton.Text = "Running..."
            
            # Clear status textbox
            $featuresStatusTextBox.Clear()
            Add-Status "Starting Windows Features Configuration..." $featuresStatusTextBox
            [System.Windows.Forms.Application]::DoEvents()
            
            # Run Windows Features Configuration
            $result = Invoke-WindowsFeaturesConfiguration -deviceType "General" -statusTextBox $featuresStatusTextBox
            
            if ($result) {
                Add-Status "Windows Features configuration completed successfully!" $featuresStatusTextBox
                $startButton.Text = "Configuration Complete"
                $startButton.BackColor = [System.Drawing.Color]::FromArgb(0, 100, 0)
            } else {
                Add-Status "Windows Features configuration failed!" $featuresStatusTextBox
                $startButton.Text = "Configuration Failed"
                $startButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
            }
            
        } catch {
            Add-Status "ERROR: $_" $featuresStatusTextBox
            $startButton.Text = "Error Occurred"
            $startButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
        }
    })
    $featuresForm.Controls.Add($startButton)

    # Close button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(320, 390)
    $closeButton.Size = New-Object System.Drawing.Size(100, 40)
    $closeButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
    $closeButton.ForeColor = [System.Drawing.Color]::White
    $closeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $closeButton.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $closeButton.FlatAppearance.BorderSize = 1
    $closeButton.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(255, 0, 0)
    $closeButton.Add_Click({
        $featuresForm.Close()
    })
    $featuresForm.Controls.Add($closeButton)

    # Add close button (X) in top-right corner
    $xCloseButton = New-Object System.Windows.Forms.Button
    $xCloseButton.Text = "✕"
    $xCloseButton.Location = New-Object System.Drawing.Point(550, 10)
    $xCloseButton.Size = New-Object System.Drawing.Size(30, 30)
    $xCloseButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $xCloseButton.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $xCloseButton.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
    $xCloseButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $xCloseButton.FlatAppearance.BorderSize = 1
    $xCloseButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $xCloseButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $xCloseButton.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $xCloseButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $xCloseButton.Add_Click({
        $featuresForm.Close()
    })
    $featuresForm.Controls.Add($xCloseButton)

    # Add KeyDown event handler for Esc key
    $featuresForm.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $featuresForm.Close()
        }
    })

    # Enable key events
    $featuresForm.KeyPreview = $true

    # When form closes, show main menu
    $featuresForm.Add_FormClosed({
        Show-MainMenu
    })

    # Show the dialog
    $featuresForm.ShowDialog()
}

# Hàm enable các features
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

# Hàm disable các features
function Invoke-DisableWindowsFeatures {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    # Lấy phiên bản hệ điều hành
    $osVersion = (Get-CimInstance Win32_OperatingSystem).Caption

    # Danh sách các features cần disable
    $featuresToDisable = @(
        @{
            Name = "Internet-Explorer-Optional-amd64"
            DisplayName = "IExplorer 11"
            Command = "dism /online /disable-feature /featurename:Internet-Explorer-Optional-amd64 /norestart"
            SupportedOS = "Windows 10"
        }
    )
    
    foreach ($feature in $featuresToDisable) {
        # Kiểm tra xem có nên thực thi trên OS hiện tại không
        if ($feature.SupportedOS -and -not ($osVersion -like "*$($feature.SupportedOS)*")) {
            Add-Status "$($feature.DisplayName): Not apply on $osVersion. Skipping..." $statusTextBox
            continue
        }
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

## STEP 6: Disk Partitioning Functions
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

# Hàm kiểm tra sự sẵn sàng của ổ đĩa
function Invoke-DiskAvailabilityCheck {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    try {
        # Lấy thông tin tất cả các ổ đĩa
        $systemDisk = Get-Disk | Where-Object { $_.IsBoot -eq $true }
        
        # Kiểm tra xem ổ hệ thống có đủ không gian để chia không
        $systemPartitions = Get-Partition -DiskNumber $systemDisk.Number
        $systemVolume = $systemPartitions | Where-Object { $_.DriveLetter -eq 'C' }
        
        if ($systemVolume) {
            $usedSpace = Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object { $_.DeviceID -eq "C:" }
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

# Hàm chọn kích thước phân vùng
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

# Hàm tạo phân vùng
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

## STEP 7: User Password Management Functions
function Invoke-UserPasswordManagement {
    param (
        [string]$deviceType,
        [System.Windows.Forms.TextBox]$statusTextBox
    )
    
    try {
        Add-Status "Starting user password management..." $statusTextBox
        
        # --- 1. Get Current User Information ---
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[-1]
        Add-Status "Current user: $currentUser" $statusTextBox
        
        # --- 2. Show Password Management Dialog ---
        $passwordResult = Show-PasswordManagementDialog -currentUser $currentUser -statusTextBox $statusTextBox
        
        if ($passwordResult) {
            Add-Status "User password management completed successfully!" $statusTextBox
        } else {
            Add-Status "User password management was cancelled or failed." $statusTextBox
        }
        
        return $true
        
    } catch {
        Add-Status "ERROR during User Password Management: $_" $statusTextBox
        return $false
    }
}

# Helper Functions cho Password Management
function Show-PasswordManagementDialog {
    param (
        [string]$currentUser,
        [System.Windows.Forms.TextBox]$statusTextBox
    )
    
    Add-Status "Opening password management dialog..." $statusTextBox
    
    try {
        # Tạo form quản lý mật khẩu
        $passwordForm = New-Object System.Windows.Forms.Form
        $passwordForm.Text = "User Password Management"
        $passwordForm.Size = New-Object System.Drawing.Size(500, 400)
        $passwordForm.StartPosition = "CenterScreen"
        $passwordForm.BackColor = [System.Drawing.Color]::Black
        $passwordForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $passwordForm.MaximizeBox = $false
        $passwordForm.MinimizeBox = $false
        
        # Gradient background
        $passwordForm.Add_Paint({
            $graphics = $_.Graphics
            $rect = New-Object System.Drawing.Rectangle(0, 0, $passwordForm.Width, $passwordForm.Height)
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
        $titleLabel.Text = "PASSWORD MANAGEMENT"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(500, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $passwordForm.Controls.Add($titleLabel)
        
        # Current user info
        $currentUser = $env:USERNAME
        $userInfoLabel = New-Object System.Windows.Forms.Label
        $userInfoLabel.Text = "Current User: $currentUser"
        $userInfoLabel.Location = New-Object System.Drawing.Point(20, 60)
        $userInfoLabel.Size = New-Object System.Drawing.Size(460, 25)
        $userInfoLabel.ForeColor = [System.Drawing.Color]::White
        $userInfoLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $userInfoLabel.BackColor = [System.Drawing.Color]::Transparent
        $passwordForm.Controls.Add($userInfoLabel)
        
        # Password input section
        $passwordLabel = New-Object System.Windows.Forms.Label
        $passwordLabel.Text = "New Password (leave empty to remove password):"
        $passwordLabel.Location = New-Object System.Drawing.Point(20, 100)
        $passwordLabel.Size = New-Object System.Drawing.Size(460, 25)
        $passwordLabel.ForeColor = [System.Drawing.Color]::Yellow
        $passwordLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $passwordLabel.BackColor = [System.Drawing.Color]::Transparent
        $passwordForm.Controls.Add($passwordLabel)
        
        # Password textbox - MẶC ĐỊNH HIỂN THỊ PASSWORD
        $passwordTextBox = New-Object System.Windows.Forms.TextBox
        $passwordTextBox.Location = New-Object System.Drawing.Point(20, 130)
        $passwordTextBox.Size = New-Object System.Drawing.Size(350, 25)
        $passwordTextBox.Font = New-Object System.Drawing.Font("Arial", 10)
        $passwordTextBox.UseSystemPasswordChar = $false  # MẶC ĐỊNH HIỂN THỊ PASSWORD
        $passwordForm.Controls.Add($passwordTextBox)
        
        # Show/Hide password checkbox - MẶC ĐỊNH CHECKED (HIỂN THỊ PASSWORD)
        $showPasswordCheckBox = New-Object System.Windows.Forms.CheckBox
        $showPasswordCheckBox.Text = "Show"
        $showPasswordCheckBox.Location = New-Object System.Drawing.Point(380, 132)
        $showPasswordCheckBox.Size = New-Object System.Drawing.Size(100, 20)
        $showPasswordCheckBox.ForeColor = [System.Drawing.Color]::White
        $showPasswordCheckBox.Font = New-Object System.Drawing.Font("Arial", 9)
        $showPasswordCheckBox.BackColor = [System.Drawing.Color]::Transparent
        $showPasswordCheckBox.Checked = $true  # MẶC ĐỊNH CHECKED (HIỂN THỊ PASSWORD)
        $showPasswordCheckBox.Add_CheckedChanged({
            # Khi CHECKED = hiển thị password, khi UNCHECKED = ẩn password
            $passwordTextBox.UseSystemPasswordChar = -not $showPasswordCheckBox.Checked
        })
        $passwordForm.Controls.Add($showPasswordCheckBox)
        
        # Instructions
        $instructionLabel = New-Object System.Windows.Forms.Label
        $instructionLabel.Text = "Enter a password to set new password`nLeave empty to remove password (no password login)"
        $instructionLabel.Location = New-Object System.Drawing.Point(20, 170)
        $instructionLabel.Size = New-Object System.Drawing.Size(460, 40)
        $instructionLabel.ForeColor = [System.Drawing.Color]::LightGray
        $instructionLabel.Font = New-Object System.Drawing.Font("Arial", 9)
        $instructionLabel.BackColor = [System.Drawing.Color]::Transparent
        $passwordForm.Controls.Add($instructionLabel)
        
        # Apply button
        $applyButton = New-Object System.Windows.Forms.Button
        $applyButton.Text = "Apply Changes"
        $applyButton.Location = New-Object System.Drawing.Point(150, 230)
        $applyButton.Size = New-Object System.Drawing.Size(120, 40)
        $applyButton.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 0)
        $applyButton.ForeColor = [System.Drawing.Color]::White
        $applyButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $applyButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $applyButton.Add_Click({
            $newPassword = $passwordTextBox.Text
            
            if ([string]::IsNullOrEmpty($newPassword)) {
                # Remove password
                $confirmResult = [System.Windows.Forms.MessageBox]::Show(
                    "Are you sure you want to remove the password for user '$currentUser'?`n`nThis will allow login without any password.",
                    "Confirm Password Removal",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )
                
                if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                    $script:passwordAction = "REMOVE"
                    $script:newPasswordValue = ""
                    $passwordForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $passwordForm.Close()
                }
            } else {
                # Set new password
                $script:passwordAction = "SET"
                $script:newPasswordValue = $newPassword
                $passwordForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $passwordForm.Close()
            }
        })
        $passwordForm.Controls.Add($applyButton)
        
        # Cancel button
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(280, 230)
        $cancelButton.Size = New-Object System.Drawing.Size(100, 40)
        $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
        $cancelButton.ForeColor = [System.Drawing.Color]::White
        $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $cancelButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $cancelButton.Add_Click({
            $passwordForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $passwordForm.Close()
        })
        $passwordForm.Controls.Add($cancelButton)
        
        # Focus vào password textbox
        $passwordForm.Add_Shown({
            $passwordTextBox.Focus()
        })
        
        # Xử lý phím Enter và ESC
        $passwordForm.KeyPreview = $true
        $passwordForm.Add_KeyDown({
            param($sender, $e)
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $applyButton.PerformClick()
            }
            elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
                $cancelButton.PerformClick()
            }
        })
        
        # Show form and get result
        $result = $passwordForm.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            if ($script:passwordAction -eq "SET") {
                Add-Status "Setting new password for user: $currentUser" $statusTextBox
                try {
                    $securePassword = ConvertTo-SecureString $script:newPasswordValue -AsPlainText -Force
                    Set-LocalUser -Name $currentUser -Password $securePassword
                    Add-Status "Password set successfully for user: $currentUser" $statusTextBox
                    return $true
                } catch {
                    Add-Status "ERROR setting password: $_" $statusTextBox
                    return $false
                }
            } elseif ($script:passwordAction -eq "REMOVE") {
                Add-Status "Removing password for user: $currentUser" $statusTextBox
                try {
                    $removeResult = Start-Process -FilePath "net" -ArgumentList "user `"$currentUser`" `"`"" -Wait -PassThru -WindowStyle Hidden
                    if ($removeResult.ExitCode -eq 0) {
                        Add-Status "Password removed successfully for user: $currentUser" $statusTextBox
                        Add-Status "User can now login without password." $statusTextBox
                        return $true
                    } else {
                        Add-Status "ERROR: Failed to remove password (Exit code: $($removeResult.ExitCode))" $statusTextBox
                        return $false
                    }
                } catch {
                    Add-Status "ERROR removing password: $_" $statusTextBox
                    return $false
                }
            }
        } else {
            Add-Status "Password management cancelled by user." $statusTextBox
            return $false
        }
        
    } catch {
        Add-Status "ERROR in password management dialog: $_" $statusTextBox
        return $false
    }
}

function Invoke-SetUserPassword {
    param (
        [string]$currentUser,
        [System.Windows.Forms.TextBox]$statusTextBox
    )
    
    Add-Status "Setting new password for user: $currentUser" $statusTextBox
    
    try {
        # Tạo form nhập mật khẩu mới
        $newPasswordForm = New-Object System.Windows.Forms.Form
        $newPasswordForm.Text = "Set New Password"
        $newPasswordForm.Size = New-Object System.Drawing.Size(450, 300)
        $newPasswordForm.StartPosition = "CenterScreen"
        $newPasswordForm.BackColor = [System.Drawing.Color]::Black
        $newPasswordForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $newPasswordForm.MaximizeBox = $false
        $newPasswordForm.MinimizeBox = $false
        
        # Gradient background
        $newPasswordForm.Add_Paint({
            $graphics = $_.Graphics
            $rect = New-Object System.Drawing.Rectangle(0, 0, $newPasswordForm.Width, $newPasswordForm.Height)
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
        $titleLabel.Text = "SET NEW PASSWORD"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(450, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $newPasswordForm.Controls.Add($titleLabel)
        
        # User info
        $userLabel = New-Object System.Windows.Forms.Label
        $userLabel.Text = "User: $currentUser"
        $userLabel.Location = New-Object System.Drawing.Point(20, 60)
        $userLabel.Size = New-Object System.Drawing.Size(400, 25)
        $userLabel.ForeColor = [System.Drawing.Color]::White
        $userLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $userLabel.BackColor = [System.Drawing.Color]::Transparent
        $newPasswordForm.Controls.Add($userLabel)
        
        # Password label
        $passwordLabel = New-Object System.Windows.Forms.Label
        $passwordLabel.Text = "New Password:"
        $passwordLabel.Location = New-Object System.Drawing.Point(20, 100)
        $passwordLabel.Size = New-Object System.Drawing.Size(120, 25)
        $passwordLabel.ForeColor = [System.Drawing.Color]::Yellow
        $passwordLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $passwordLabel.BackColor = [System.Drawing.Color]::Transparent
        $newPasswordForm.Controls.Add($passwordLabel)
        
        # Password textbox
        $passwordTextBox = New-Object System.Windows.Forms.TextBox
        $passwordTextBox.Location = New-Object System.Drawing.Point(150, 98)
        $passwordTextBox.Size = New-Object System.Drawing.Size(250, 25)
        $passwordTextBox.Font = New-Object System.Drawing.Font("Arial", 10)
        $passwordTextBox.UseSystemPasswordChar = $true
        $newPasswordForm.Controls.Add($passwordTextBox)
        
        # Confirm password label
        $confirmLabel = New-Object System.Windows.Forms.Label
        $confirmLabel.Text = "Confirm Password:"
        $confirmLabel.Location = New-Object System.Drawing.Point(20, 140)
        $confirmLabel.Size = New-Object System.Drawing.Size(120, 25)
        $confirmLabel.ForeColor = [System.Drawing.Color]::Yellow
        $confirmLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $confirmLabel.BackColor = [System.Drawing.Color]::Transparent
        $newPasswordForm.Controls.Add($confirmLabel)
        
        # Confirm password textbox
        $confirmTextBox = New-Object System.Windows.Forms.TextBox
        $confirmTextBox.Location = New-Object System.Drawing.Point(150, 138)
        $confirmTextBox.Size = New-Object System.Drawing.Size(250, 25)
        $confirmTextBox.Font = New-Object System.Drawing.Font("Arial", 10)
        $confirmTextBox.UseSystemPasswordChar = $true
        $newPasswordForm.Controls.Add($confirmTextBox)
        
        # OK button
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "Set Password"
        $okButton.Location = New-Object System.Drawing.Point(150, 190)
        $okButton.Size = New-Object System.Drawing.Size(100, 35)
        $okButton.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 0)
        $okButton.ForeColor = [System.Drawing.Color]::White
        $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $okButton.Add_Click({
            if ($passwordTextBox.Text -eq $confirmTextBox.Text) {
                if ($passwordTextBox.Text.Length -ge 1) {
                    $script:newPassword = $passwordTextBox.Text
                    $newPasswordForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $newPasswordForm.Close()
                } else {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Password must be at least 1 character long",
                        "Invalid Password",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    )
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Passwords do not match",
                    "Password Mismatch",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }
        })
        $newPasswordForm.Controls.Add($okButton)
        
        # Cancel button
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(260, 190)
        $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
        $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
        $cancelButton.ForeColor = [System.Drawing.Color]::White
        $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $cancelButton.Add_Click({
            $newPasswordForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $newPasswordForm.Close()
        })
        $newPasswordForm.Controls.Add($cancelButton)
        
        # Focus vào password textbox
        $newPasswordForm.Add_Shown({
            $passwordTextBox.Focus()
        })
        
        # Show form and get result
        $result = $newPasswordForm.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            # Set password using PowerShell
            try {
                $securePassword = ConvertTo-SecureString $script:newPassword -AsPlainText -Force
                Set-LocalUser -Name $currentUser -Password $securePassword
                Add-Status "Password set successfully for user: $currentUser" $statusTextBox
                return $true
            } catch {
                Add-Status "ERROR setting password: $_" $statusTextBox
                return $false
            }
        } else {
            Add-Status "Set password cancelled by user." $statusTextBox
            return $false
        }
        
    } catch {
        Add-Status "ERROR in set password function: $_" $statusTextBox
        return $false
    }
}

function Invoke-RemoveUserPassword {
    param (
        [string]$currentUser,
        [System.Windows.Forms.TextBox]$statusTextBox
    )
    
    Add-Status "Removing password for user: $currentUser" $statusTextBox
    
    try {
        # Confirmation dialog
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to remove the password for user '$currentUser'?`n`nThis will allow login without any password.",
            "Confirm Password Removal",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Remove password using net user command
            try {
                $removeResult = Start-Process -FilePath "net" -ArgumentList "user `"$currentUser`" `"`"" -Wait -PassThru -WindowStyle Hidden
                if ($removeResult.ExitCode -eq 0) {
                    Add-Status "Password removed successfully for user: $currentUser" $statusTextBox
                    Add-Status "User can now login without password." $statusTextBox
                    return $true
                } else {
                    Add-Status "ERROR: Failed to remove password (Exit code: $($removeResult.ExitCode))" $statusTextBox
                    return $false
                }
            } catch {
                Add-Status "ERROR removing password: $_" $statusTextBox
                return $false
            }
        } else {
            Add-Status "Password removal cancelled by user." $statusTextBox
            return $false
        }
        
    } catch {
        Add-Status "ERROR in remove password function: $_" $statusTextBox
        return $false
    }
}

## STEP 8: Domain Management Functions
function Invoke-DomainJoin {
    param (
        [string]$deviceType,
        [System.Windows.Forms.TextBox]$statusTextBox
    )
    
    try {
        Add-Status "Starting domain join process..." $statusTextBox
        
        # --- 1. Check Current Domain Status ---
        $currentDomain = Invoke-CheckDomainStatus $statusTextBox
        
        if ($currentDomain) {
            Add-Status "Computer is already joined to domain: $currentDomain" $statusTextBox
            Add-Status "Domain join skipped." $statusTextBox
            return $true
        }
        
        # --- 2. Show Domain Join Dialog ---
        $domainResult = Show-DomainJoinDialog -statusTextBox $statusTextBox
        
        if ($domainResult) {
            Add-Status "Domain join completed successfully!" $statusTextBox
        } else {
            Add-Status "Domain join was cancelled or failed." $statusTextBox
        }
        
        return $true
        
    } catch {
        Add-Status "ERROR during Domain Join: $_" $statusTextBox
        return $false
    }
}

# Helper Functions cho Domain Join
function Invoke-CheckDomainStatus {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    try {
        Add-Status "Checking current domain status..." $statusTextBox
        
        # Kiểm tra xem máy tính đã join domain chưa
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        
        if ($computerSystem.PartOfDomain) {
            $domainName = $computerSystem.Domain
            Add-Status "Computer is currently joined to domain: $domainName" $statusTextBox
            return $domainName
        } else {
            Add-Status "Computer is currently in workgroup: $($computerSystem.Workgroup)" $statusTextBox
            return $null
        }
        
    } catch {
        Add-Status "Warning: Could not check domain status: $_" $statusTextBox
        return $null
    }
}

#
function Show-DomainJoinDialog {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    Add-Status "Opening domain management dialog..." $statusTextBox
    
    try {
        # Tạo form quản lý domain
        $domainForm = New-Object System.Windows.Forms.Form
        $domainForm.Text = "Domain Management"
        $domainForm.Size = New-Object System.Drawing.Size(500, 500)
        $domainForm.StartPosition = "CenterScreen"
        $domainForm.BackColor = [System.Drawing.Color]::Black
        $domainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $domainForm.MaximizeBox = $false
        $domainForm.MinimizeBox = $false
        
        # Gradient background
        $domainForm.Add_Paint({
            $graphics = $_.Graphics
            $rect = New-Object System.Drawing.Rectangle(0, 0, $domainForm.Width, $domainForm.Height)
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
        $titleLabel.Text = "DOMAIN MANAGEMENT"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(500, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $domainForm.Controls.Add($titleLabel)
        
        # Current status info
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        $currentStatus = if ($computerSystem.PartOfDomain) {
            "Domain: $($computerSystem.Domain)"
        } else {
            "Workgroup: $($computerSystem.Workgroup)"
        }
        
        $statusLabel = New-Object System.Windows.Forms.Label
        $statusLabel.Text = "Current Status: $currentStatus"
        $statusLabel.Location = New-Object System.Drawing.Point(20, 60)
        $statusLabel.Size = New-Object System.Drawing.Size(460, 25)
        $statusLabel.ForeColor = [System.Drawing.Color]::White
        $statusLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $statusLabel.BackColor = [System.Drawing.Color]::Transparent
        $domainForm.Controls.Add($statusLabel)
        
        # Variable to store selected action
        $script:selectedDomainAction = ""
        
        # JOIN DOMAIN Button
        $btnJoinDomain = New-Object System.Windows.Forms.Button
        $btnJoinDomain.Text = "JOIN DOMAIN"
        $btnJoinDomain.Location = New-Object System.Drawing.Point(50, 100)
        $btnJoinDomain.Size = New-Object System.Drawing.Size(400, 50)
        $btnJoinDomain.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 0)
        $btnJoinDomain.ForeColor = [System.Drawing.Color]::White
        $btnJoinDomain.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnJoinDomain.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
        $btnJoinDomain.Add_Click({
            $script:selectedDomainAction = "JOIN"
            $domainForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $domainForm.Close()
        })
        $domainForm.Controls.Add($btnJoinDomain)
        
        # LEAVE DOMAIN Button
        $btnLeaveDomain = New-Object System.Windows.Forms.Button
        $btnLeaveDomain.Text = "LEAVE DOMAIN (Join Workgroup)"
        $btnLeaveDomain.Location = New-Object System.Drawing.Point(50, 160)
        $btnLeaveDomain.Size = New-Object System.Drawing.Size(400, 50)
        $btnLeaveDomain.BackColor = [System.Drawing.Color]::FromArgb(150, 100, 0)
        $btnLeaveDomain.ForeColor = [System.Drawing.Color]::White
        $btnLeaveDomain.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnLeaveDomain.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
        $btnLeaveDomain.Add_Click({
            $script:selectedDomainAction = "LEAVE"
            $domainForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $domainForm.Close()
        })
        $domainForm.Controls.Add($btnLeaveDomain)
        
        # JOIN WORKGROUP Button
        $btnJoinWorkgroup = New-Object System.Windows.Forms.Button
        $btnJoinWorkgroup.Text = "JOIN WORKGROUP"
        $btnJoinWorkgroup.Location = New-Object System.Drawing.Point(50, 220)
        $btnJoinWorkgroup.Size = New-Object System.Drawing.Size(400, 50)
        $btnJoinWorkgroup.BackColor = [System.Drawing.Color]::FromArgb(0, 100, 150)
        $btnJoinWorkgroup.ForeColor = [System.Drawing.Color]::White
        $btnJoinWorkgroup.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnJoinWorkgroup.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
        $btnJoinWorkgroup.Add_Click({
            $script:selectedDomainAction = "WORKGROUP"
            $domainForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $domainForm.Close()
        })
        $domainForm.Controls.Add($btnJoinWorkgroup)
        
        # Instructions
        $instructionLabel = New-Object System.Windows.Forms.Label
        $instructionLabel.Text = "• Join Domain: Connect to vietunion.local domain`n• Leave Domain: Disconnect from current domain and join workgroup`n• Join Workgroup: Connect to a specific workgroup"
        $instructionLabel.Location = New-Object System.Drawing.Point(20, 290)
        $instructionLabel.Size = New-Object System.Drawing.Size(460, 60)
        $instructionLabel.ForeColor = [System.Drawing.Color]::LightGray
        $instructionLabel.Font = New-Object System.Drawing.Font("Arial", 9)
        $instructionLabel.BackColor = [System.Drawing.Color]::Transparent
        $domainForm.Controls.Add($instructionLabel)
        
        # Cancel button
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(200, 370)
        $cancelButton.Size = New-Object System.Drawing.Size(100, 40)
        $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
        $cancelButton.ForeColor = [System.Drawing.Color]::White
        $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $cancelButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $cancelButton.Add_Click({
            $domainForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $domainForm.Close()
        })
        $domainForm.Controls.Add($cancelButton)
        
        # Show form and get result
        $result = $domainForm.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            if ($script:selectedDomainAction -eq "JOIN") {
                return Invoke-JoinDomainProcess $statusTextBox
            } elseif ($script:selectedDomainAction -eq "LEAVE") {
                return Invoke-LeaveDomainProcess $statusTextBox
            } elseif ($script:selectedDomainAction -eq "WORKGROUP") {
                return Invoke-JoinWorkgroupProcess $statusTextBox
            }
        } else {
            Add-Status "Domain management cancelled by user." $statusTextBox
            return $false
        }
        
    } catch {
        Add-Status "ERROR in domain management dialog: $_" $statusTextBox
        return $false
    }
}

# Hàm xử lý Join Domain
function Invoke-JoinDomainProcess {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    try {
        # Tạo form nhập thông tin domain
        $joinForm = New-Object System.Windows.Forms.Form
        $joinForm.Text = "Join Domain"
        $joinForm.Size = New-Object System.Drawing.Size(450, 300)
        $joinForm.StartPosition = "CenterScreen"
        $joinForm.BackColor = [System.Drawing.Color]::Black
        $joinForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $joinForm.MaximizeBox = $false
        $joinForm.MinimizeBox = $false
        
        # Gradient background
        $joinForm.Add_Paint({
            $graphics = $_.Graphics
            $rect = New-Object System.Drawing.Rectangle(0, 0, $joinForm.Width, $joinForm.Height)
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
        $titleLabel.Text = "JOIN DOMAIN"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(450, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $joinForm.Controls.Add($titleLabel)
        
        # Domain name (pre-filled)
        $domainLabel = New-Object System.Windows.Forms.Label
        $domainLabel.Text = "Domain: vietunion.local"
        $domainLabel.Location = New-Object System.Drawing.Point(20, 70)
        $domainLabel.Size = New-Object System.Drawing.Size(400, 25)
        $domainLabel.ForeColor = [System.Drawing.Color]::Yellow
        $domainLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $domainLabel.BackColor = [System.Drawing.Color]::Transparent
        $joinForm.Controls.Add($domainLabel)
        
        # Username
        $usernameLabel = New-Object System.Windows.Forms.Label
        $usernameLabel.Text = "Username:"
        $usernameLabel.Location = New-Object System.Drawing.Point(20, 110)
        $usernameLabel.Size = New-Object System.Drawing.Size(80, 25)
        $usernameLabel.ForeColor = [System.Drawing.Color]::White
        $usernameLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $usernameLabel.BackColor = [System.Drawing.Color]::Transparent
        $joinForm.Controls.Add($usernameLabel)
        
        $usernameTextBox = New-Object System.Windows.Forms.TextBox
        $usernameTextBox.Location = New-Object System.Drawing.Point(110, 108)
        $usernameTextBox.Size = New-Object System.Drawing.Size(250, 25)
        $usernameTextBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $usernameTextBox.Text = "-HDK-hieudang"
        $joinForm.Controls.Add($usernameTextBox)
        
        # Password
        $passwordLabel = New-Object System.Windows.Forms.Label
        $passwordLabel.Text = "Password:"
        $passwordLabel.Location = New-Object System.Drawing.Point(20, 150)
        $passwordLabel.Size = New-Object System.Drawing.Size(80, 25)
        $passwordLabel.ForeColor = [System.Drawing.Color]::White
        $passwordLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $passwordLabel.BackColor = [System.Drawing.Color]::Transparent
        $joinForm.Controls.Add($passwordLabel)
        
        $passwordTextBox = New-Object System.Windows.Forms.TextBox
        $passwordTextBox.Location = New-Object System.Drawing.Point(110, 148)
        $passwordTextBox.Size = New-Object System.Drawing.Size(250, 25)
        $passwordTextBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $passwordTextBox.UseSystemPasswordChar = $true
        $joinForm.Controls.Add($passwordTextBox)
        
        # Join button
        $joinButton = New-Object System.Windows.Forms.Button
        $joinButton.Text = "Join Domain"
        $joinButton.Location = New-Object System.Drawing.Point(150, 200)
        $joinButton.Size = New-Object System.Drawing.Size(100, 35)
        $joinButton.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 0)
        $joinButton.ForeColor = [System.Drawing.Color]::White
        $joinButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $joinButton.Add_Click({
            if ($usernameTextBox.Text.Trim() -and $passwordTextBox.Text) {
                $script:domainUsername = $usernameTextBox.Text.Trim()
                $script:domainPassword = $passwordTextBox.Text
                $joinForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $joinForm.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Please enter username and password.", "Missing Information")
            }
        })
        $joinForm.Controls.Add($joinButton)
        
        # Cancel button
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(260, 200)
        $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
        $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
        $cancelButton.ForeColor = [System.Drawing.Color]::White
        $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $cancelButton.Add_Click({
            $joinForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $joinForm.Close()
        })
        $joinForm.Controls.Add($cancelButton)
        
        $result = $joinForm.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            Add-Status "Joining domain vietunion.local..." $statusTextBox
            try {
                $securePassword = ConvertTo-SecureString $script:domainPassword -AsPlainText -Force
                $credential = New-Object System.Management.Automation.PSCredential("vietunion.local\$($script:domainUsername)", $securePassword)
                
                Add-Computer -DomainName "vietunion.local" -Credential $credential -Restart -Force
                Add-Status "Successfully joined domain! Computer will restart." $statusTextBox
                return $true
            } catch {
                Add-Status "ERROR joining domain: $_" $statusTextBox
                return $false
            }
        }
        return $false
        
    } catch {
        Add-Status "ERROR in join domain process: $_" $statusTextBox
        return $false
    }
}

# Hàm xử lý Leave Domain
function Invoke-LeaveDomainProcess {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    try {
        # Kiểm tra xem máy có đang trong domain không
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        if (-not $computerSystem.PartOfDomain) {
            Add-Status "Computer is not currently joined to any domain." $statusTextBox
            [System.Windows.Forms.MessageBox]::Show("Computer is not currently joined to any domain.", "Not in Domain")
            return $false
        }
        
        # Tạo form nhập credentials để leave domain
        $leaveForm = New-Object System.Windows.Forms.Form
        $leaveForm.Text = "Leave Domain"
        $leaveForm.Size = New-Object System.Drawing.Size(450, 300)
        $leaveForm.StartPosition = "CenterScreen"
        $leaveForm.BackColor = [System.Drawing.Color]::Black
        $leaveForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $leaveForm.MaximizeBox = $false
        $leaveForm.MinimizeBox = $false
        
        # Gradient background
        $leaveForm.Add_Paint({
            $graphics = $_.Graphics
            $rect = New-Object System.Drawing.Rectangle(0, 0, $leaveForm.Width, $leaveForm.Height)
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
        $titleLabel.Text = "LEAVE DOMAIN"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(450, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $leaveForm.Controls.Add($titleLabel)
        
        # Current domain info
        $currentDomainLabel = New-Object System.Windows.Forms.Label
        $currentDomainLabel.Text = "Current Domain: $($computerSystem.Domain)"
        $currentDomainLabel.Location = New-Object System.Drawing.Point(20, 70)
        $currentDomainLabel.Size = New-Object System.Drawing.Size(400, 25)
        $currentDomainLabel.ForeColor = [System.Drawing.Color]::Yellow
        $currentDomainLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
        $currentDomainLabel.BackColor = [System.Drawing.Color]::Transparent
        $leaveForm.Controls.Add($currentDomainLabel)
        
        # Username
        $usernameLabel = New-Object System.Windows.Forms.Label
        $usernameLabel.Text = "Domain Admin Username:"
        $usernameLabel.Location = New-Object System.Drawing.Point(20, 110)
        $usernameLabel.Size = New-Object System.Drawing.Size(150, 25)
        $usernameLabel.ForeColor = [System.Drawing.Color]::White
        $usernameLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $usernameLabel.BackColor = [System.Drawing.Color]::Transparent
        $leaveForm.Controls.Add($usernameLabel)
        
        $usernameTextBox = New-Object System.Windows.Forms.TextBox
        $usernameTextBox.Location = New-Object System.Drawing.Point(180, 108)
        $usernameTextBox.Size = New-Object System.Drawing.Size(180, 25)
        $usernameTextBox.Font = New-Object System.Drawing.Font("Arial", 10)
        $usernameTextBox.Text = "-HDK-hieudang"
        $leaveForm.Controls.Add($usernameTextBox)
        
        # Password
        $passwordLabel = New-Object System.Windows.Forms.Label
        $passwordLabel.Text = "Password:"
        $passwordLabel.Location = New-Object System.Drawing.Point(20, 150)
        $passwordLabel.Size = New-Object System.Drawing.Size(150, 25)
        $passwordLabel.ForeColor = [System.Drawing.Color]::White
        $passwordLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $passwordLabel.BackColor = [System.Drawing.Color]::Transparent
        $leaveForm.Controls.Add($passwordLabel)
        
        $passwordTextBox = New-Object System.Windows.Forms.TextBox
        $passwordTextBox.Location = New-Object System.Drawing.Point(180, 148)
        $passwordTextBox.Size = New-Object System.Drawing.Size(180, 25)
        $passwordTextBox.Font = New-Object System.Drawing.Font("Arial", 10)
        $passwordTextBox.UseSystemPasswordChar = $true
        $leaveForm.Controls.Add($passwordTextBox)
        
        # Leave button
        $leaveButton = New-Object System.Windows.Forms.Button
        $leaveButton.Text = "Leave Domain"
        $leaveButton.Location = New-Object System.Drawing.Point(150, 200)
        $leaveButton.Size = New-Object System.Drawing.Size(100, 35)
        $leaveButton.BackColor = [System.Drawing.Color]::FromArgb(150, 100, 0)
        $leaveButton.ForeColor = [System.Drawing.Color]::White
        $leaveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $leaveButton.Add_Click({
            if ($usernameTextBox.Text.Trim() -and $passwordTextBox.Text) {
                $script:domainUsername = $usernameTextBox.Text.Trim()
                $script:domainPassword = $passwordTextBox.Text
                $leaveForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $leaveForm.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Please enter username and password.", "Missing Information")
            }
        })
        $leaveForm.Controls.Add($leaveButton)
        
        # Cancel button
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(260, 200)
        $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
        $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
        $cancelButton.ForeColor = [System.Drawing.Color]::White
        $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $cancelButton.Add_Click({
            $leaveForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $leaveForm.Close()
        })
        $leaveForm.Controls.Add($cancelButton)
        
        $result = $leaveForm.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            Add-Status "Leaving domain $($computerSystem.Domain)..." $statusTextBox
            try {
                $securePassword = ConvertTo-SecureString $script:domainPassword -AsPlainText -Force
                $credential = New-Object System.Management.Automation.PSCredential("$($computerSystem.Domain)\$($script:domainUsername)", $securePassword)
                
                Remove-Computer -UnjoinDomainCredential $credential -WorkgroupName "WORKGROUP" -Restart -Force
                Add-Status "Successfully left domain! Computer will restart and join WORKGROUP." $statusTextBox
                return $true
            } catch {
                Add-Status "ERROR leaving domain: $_" $statusTextBox
                return $false
            }
        }
        return $false
        
    } catch {
        Add-Status "ERROR in leave domain process: $_" $statusTextBox
        return $false
    }
}

# Hàm xử lý Join Workgroup
function Invoke-JoinWorkgroupProcess {
    param ([System.Windows.Forms.TextBox]$statusTextBox)
    
    try {
        # Tạo form nhập tên workgroup
        $workgroupForm = New-Object System.Windows.Forms.Form
        $workgroupForm.Text = "Join Workgroup"
        $workgroupForm.Size = New-Object System.Drawing.Size(400, 200)
        $workgroupForm.StartPosition = "CenterScreen"
        $workgroupForm.BackColor = [System.Drawing.Color]::Black
        $workgroupForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $workgroupForm.MaximizeBox = $false
        $workgroupForm.MinimizeBox = $false
        
        # Gradient background
        $workgroupForm.Add_Paint({
            $graphics = $_.Graphics
            $rect = New-Object System.Drawing.Rectangle(0, 0, $workgroupForm.Width, $workgroupForm.Height)
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
        $titleLabel.Text = "JOIN WORKGROUP"
        $titleLabel.Location = New-Object System.Drawing.Point(0, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(400, 30)
        $titleLabel.ForeColor = [System.Drawing.Color]::Lime
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.BackColor = [System.Drawing.Color]::Transparent
        $workgroupForm.Controls.Add($titleLabel)
        
        # Workgroup name
        $workgroupLabel = New-Object System.Windows.Forms.Label
        $workgroupLabel.Text = "Workgroup Name:"
        $workgroupLabel.Location = New-Object System.Drawing.Point(20, 70)
        $workgroupLabel.Size = New-Object System.Drawing.Size(120, 25)
        $workgroupLabel.ForeColor = [System.Drawing.Color]::White
        $workgroupLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $workgroupLabel.BackColor = [System.Drawing.Color]::Transparent
        $workgroupForm.Controls.Add($workgroupLabel)
        
        $workgroupTextBox = New-Object System.Windows.Forms.TextBox
        $workgroupTextBox.Location = New-Object System.Drawing.Point(150, 68)
        $workgroupTextBox.Size = New-Object System.Drawing.Size(200, 25)
        $workgroupTextBox.Font = New-Object System.Drawing.Font("Arial", 10)
        $workgroupTextBox.Text = "WORKGROUP"
        $workgroupForm.Controls.Add($workgroupTextBox)
        
        # Join button
        $joinButton = New-Object System.Windows.Forms.Button
        $joinButton.Text = "Join"
        $joinButton.Location = New-Object System.Drawing.Point(120, 120)
        $joinButton.Size = New-Object System.Drawing.Size(80, 35)
        $joinButton.BackColor = [System.Drawing.Color]::FromArgb(0, 100, 150)
        $joinButton.ForeColor = [System.Drawing.Color]::White
        $joinButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $joinButton.Add_Click({
            if ($workgroupTextBox.Text.Trim()) {
                $script:workgroupName = $workgroupTextBox.Text.Trim()
                $workgroupForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $workgroupForm.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Please enter workgroup name.", "Missing Information")
            }
        })
        $workgroupForm.Controls.Add($joinButton)
        
        # Cancel button
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(210, 120)
        $cancelButton.Size = New-Object System.Drawing.Size(80, 35)
        $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
        $cancelButton.ForeColor = [System.Drawing.Color]::White
        $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $cancelButton.Add_Click({
            $workgroupForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $workgroupForm.Close()
        })
        $workgroupForm.Controls.Add($cancelButton)
        
        $result = $workgroupForm.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            Add-Status "Joining workgroup $($script:workgroupName)..." $statusTextBox
            try {
                # Kiểm tra xem có đang trong domain không
                $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
                if ($computerSystem.PartOfDomain) {
                    Add-Status "ERROR: Computer is currently in domain. Please leave domain first." $statusTextBox
                    return $false
                }
                
                Add-Computer -WorkgroupName $script:workgroupName -Restart -Force
                Add-Status "Successfully joined workgroup $($script:workgroupName)! Computer will restart." $statusTextBox
                return $true
            } catch {
                Add-Status "ERROR joining workgroup: $_" $statusTextBox
                return $false
            }
        }
        return $false
        
    } catch {
        Add-Status "ERROR in join workgroup process: $_" $statusTextBox
        return $false
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
    $statusForm.Size = New-Object System.Drawing.Size(595, 480)
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
    $progressBar.Size = New-Object System.Drawing.Size(560, 15)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $statusForm.Controls.Add($progressBar)

    # Show the form
    $statusForm.Show()
    [System.Windows.Forms.Application]::DoEvents()


    
    try {
        Add-Status "Pre-Step: Connecting to WiFi network..." $statusTextBox
        $progressBar.Value = 5
        
        $wifiResult = Invoke-WiFiAutoConnection $statusTextBox
        if ($wifiResult) {
            Add-Status "WiFi connection completed!" $statusTextBox
        } else {
            Add-Status "WiFi connection failed, but continuing..." $statusTextBox
        }

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
        $copyResult = Copy-SoftwareFiles -deviceType $deviceType $statusTextBox
        if (-not $copyResult) {
            Add-Status "Error copying software files. Exiting..." $statusTextBox
            return
        }

        # Install software (gọi hàm toàn cục)
        Add-Status "Installing software..." $statusTextBox
        Install-Software -deviceType $deviceType $statusTextBox
        Add-Status "All installation completed successfully for $deviceType" $statusTextBox
        Add-Status "STEP 1 completed successfully!" $statusTextBox

        # STEP 2: System Configuration and Shortcut Creation
        Add-Status "STEP 2: Configuring System and Creating Shortcuts..." $statusTextBox
        $progressBar.Value = 28 # Tăng giá trị progress bar

        $configResult = Invoke-SystemConfiguration -deviceType $deviceType -statusTextBox $statusTextBox
        if ($configResult) {
            Add-Status "STEP 2 completed successfully!" $statusTextBox
        } else {
            Add-Status "STEP 2 encountered errors. Check logs." $statusTextBox
        }

        # STEP 3: System Cleanup and Optimization
        Add-Status "STEP 3: Cleaning up system and optimizing performance..." $statusTextBox
        $progressBar.Value = 42 # Tăng giá trị progress bar

        $cleanupResult = Invoke-SystemCleanup -deviceType $deviceType -statusTextBox $statusTextBox
        if ($cleanupResult) {
            Add-Status "STEP 3 completed successfully!" $statusTextBox
        } else {
            Add-Status "STEP 3 encountered errors. Check logs." $statusTextBox
        }

        # STEP 4: Windows and Office Activation
        Add-Status "STEP 4: Activating Windows 10 Pro and Office 2019 Pro Plus..." $statusTextBox
        $progressBar.Value = 56 # Tăng giá trị progress bar

        $activationResult = Invoke-ActivationConfiguration -deviceType $deviceType -statusTextBox $statusTextBox
        if ($activationResult) {
            Add-Status "STEP 4 completed successfully!" $statusTextBox
        } else {
            Add-Status "STEP 4 encountered errors. Check logs." $statusTextBox
        }

        # STEP 5: Windows Features Configuration
        Add-Status "STEP 5: Configuring Windows Features..." $statusTextBox
        $progressBar.Value = 70 # Tăng giá trị progress bar

        $featuresResult = Invoke-WindowsFeaturesConfiguration -deviceType $deviceType -statusTextBox $statusTextBox
        if ($featuresResult) {
            Add-Status "STEP 5 completed successfully!" $statusTextBox
        } else {
            Add-Status "STEP 5 encountered errors. Check logs." $statusTextBox
        }

        # STEP 6: Disk Partitioning (Laptop only)
        if ($deviceType -eq "Desktop") {
            Add-Status "STEP 6: Skipped for Desktop device type" $statusTextBox
            $progressBar.Value = 85
        } else {
            Add-Status "STEP 6: Configuring disk partitioning..." $statusTextBox
            $progressBar.Value = 85
        
            $partitioningResult = Invoke-DiskPartitioning -deviceType $deviceType -statusTextBox $statusTextBox
            if ($partitioningResult) {
                Add-Status "STEP 6 completed successfully!" $statusTextBox
            } else {
                Add-Status "STEP 6 encountered errors. Check logs." $statusTextBox
            }
        }

        # STEP 7: User Password Management
        Add-Status "STEP 7: Managing user password..." $statusTextBox
        $progressBar.Value = 95

        $passwordResult = Invoke-UserPasswordManagement -deviceType $deviceType -statusTextBox $statusTextBox
        if ($passwordResult) {
            Add-Status "STEP 7 completed successfully!" $statusTextBox
        } else {
            Add-Status "STEP 7 encountered errors. Check logs." $statusTextBox
        }

        # Step 8: Domain Join
        Add-Status "Step 8/8: Joining domain..." $statusTextBox
        $progressBar.Value = 100

        $domainResult = Invoke-JoinDomainProcess -deviceType $deviceType -statusTextBox $statusTextBox
        if ($domainResult) {
            Add-Status "Step 8 completed successfully!" $statusTextBox
        } else {
            Add-Status "Step 8 encountered errors. Check logs." $statusTextBox
        }

        Add-Status "All steps completed! System setup finished." $statusTextBox
        Add-Status "Computer will restart if domain join was successful." $statusTextBox
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

            Add-Status "Shrinking drive $driveLetter..."

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
    Show-TurnOnFeaturesDialog
}
# [7] Rename Device
$buttonRenameDevice = New-DynamicButton -text "[7] Rename Device" -x 430 -y 180 -width 380 -height 60 -clickAction {
    Show-RenameDeviceDialog
}
# [8] Set Password
$buttonSetPassword = New-DynamicButton -text "[8] Set Password" -x 430 -y 260 -width 380 -height 60 -clickAction {
    Show-PasswordManagementDialog
}
# [9] Join Domain
$buttonJoinDomain = New-DynamicButton -text "[9] Join Domain" -x 430 -y 340 -width 380 -height 60 -clickAction {
    Show-DomainJoinDialog
}
# [0] Exit
$buttonExit = New-DynamicButton -text "[0] Exit" -x 430 -y 420 -width 380 -height 60 -normalColor ([System.Drawing.Color]::FromArgb(180, 0, 0)) -hoverColor ([System.Drawing.Color]::FromArgb(220, 0, 0)) -pressColor ([System.Drawing.Color]::FromArgb(120, 0, 0)) -clickAction {
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
# Start Application
$script:form.ShowDialog() 