Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "BAOPROVIP - Hệ thống quản lý"
$form.Size = New-Object System.Drawing.Size(850, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::Black

# Tạo tiêu đề
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "WELCOME TO BAOPROVIP"
$titleLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0) # Green color
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$titleLabel.Size = New-Object System.Drawing.Size($form.ClientSize.Width, 50)
$titleLabel.Location = New-Object System.Drawing.Point(0, 20)
$form.Controls.Add($titleLabel)

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

# Left column buttons
$buttonRunAll = New-GreenButton -text "Run All Options" -x 30 -y 100 -width 380 -height 60 -clickAction {
    [System.Windows.Forms.MessageBox]::Show("Running all options...")
    # Add commands to run here
}

$buttonInstallSoftware = New-GreenButton -text "Install All Software" -x 30 -y 180 -width 380 -height 60 -clickAction {
    [System.Windows.Forms.MessageBox]::Show("Installing all software...")
    # Add installation commands here
}

$buttonPowerOptions = New-GreenButton -text "Power Options and Firewall" -x 30 -y 260 -width 380 -height 60 -clickAction {
    [System.Windows.Forms.MessageBox]::Show("Configuring power options and firewall...")
    # Add power and firewall configuration commands here
}

$buttonChangeVolume = New-GreenButton -text "Change / Edit Volume" -x 30 -y 340 -width 380 -height 60 -clickAction {
    [System.Windows.Forms.MessageBox]::Show("Opening volume management...")
    # Add volume management commands here
}

$buttonActivate = New-GreenButton -text "Activate" -x 30 -y 420 -width 380 -height 60 -clickAction {
    [System.Windows.Forms.MessageBox]::Show("Opening activation options...")
    # Add activation commands here
}

# Right column buttons
$buttonTurnOnFeatures = New-GreenButton -text "Turn On Features" -x 430 -y 100 -width 380 -height 60 -clickAction {
    # Create Windows Features form
    $featuresForm = New-Object System.Windows.Forms.Form
    $featuresForm.Text = "Windows Features"
    $featuresForm.Size = New-Object System.Drawing.Size(500, 500)
    $featuresForm.StartPosition = "CenterScreen"
    $featuresForm.BackColor = [System.Drawing.Color]::Black
    $featuresForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $featuresForm.MaximizeBox = $false
    $featuresForm.MinimizeBox = $false

    # Title label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "WINDOWS FEATURES MANAGEMENT"
    $titleLabel.Location = New-Object System.Drawing.Point(50, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(400, 30)
    $titleLabel.ForeColor = [System.Drawing.Color]::Lime
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $featuresForm.Controls.Add($titleLabel)

    # Status text box
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Multiline = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.Location = New-Object System.Drawing.Point(50, 320)
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

    # Enable .NET Framework 3.5 button
    $btnEnableNetFx = New-GreenButton -text "Enable .NET Framework 3.5" -x 50 -y 80 -width 400 -height 40 -clickAction {
        Add-Status "Checking .NET Framework 3.5 status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:NetFx3"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Disabled") {
                Add-Status "Enabling .NET Framework 3.5..."
                $enableCmd = "dism /online /enable-feature /featurename:NetFx3 /all /norestart"
                $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $enableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -PassThru -Wait
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
    $btnEnableWcfHttp = New-GreenButton -text "Enable WCF-HTTP-Activation" -x 50 -y 130 -width 400 -height 40 -clickAction {
        Add-Status "Checking WCF-HTTP-Activation status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:WCF-HTTP-Activation"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Disabled") {
                Add-Status "Enabling WCF-HTTP-Activation..."
                $enableCmd = "DISM /Online /Enable-Feature /FeatureName:WCF-HTTP-Activation /All /Quiet /NoRestart"
                $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $enableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -PassThru -Wait
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
    $btnEnableWcfNonHttp = New-GreenButton -text "Enable WCF-NonHTTP-Activation" -x 50 -y 180 -width 400 -height 40 -clickAction {
        Add-Status "Checking WCF-NonHTTP-Activation status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:WCF-NonHTTP-Activation"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Disabled") {
                Add-Status "Enabling WCF-NonHTTP-Activation..."
                $enableCmd = "DISM /Online /Enable-Feature /FeatureName:WCF-NonHTTP-Activation /All /Quiet /NoRestart"
                $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $enableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -PassThru -Wait
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
    $btnDisableIE = New-GreenButton -text "Disable Internet Explorer 11" -x 50 -y 230 -width 400 -height 40 -clickAction {
        Add-Status "Checking Internet Explorer 11 status..."
        try {
            $command = "dism /online /get-featureinfo /featurename:Internet-Explorer-Optional-amd64"
            $output = Invoke-Expression $command | Out-String

            if ($output -match "State : Enabled") {
                Add-Status "Disabling Internet Explorer 11..."
                $disableCmd = "dism /online /disable-feature /featurename:Internet-Explorer-Optional-amd64 /norestart"
                $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command Start-Process cmd.exe -ArgumentList '/c $disableCmd' -Verb RunAs -WindowStyle Hidden -Wait" -PassThru -Wait
                Add-Status "Internet Explorer 11 has been disabled."
            } else {
                Add-Status "Internet Explorer 11 is already disabled."
            }
        } catch {
            Add-Status "Error: $_"
        }
    }
    $featuresForm.Controls.Add($btnDisableIE)

    # Return to Main Menu button (màu đỏ)
    $btnReturn = New-Object System.Windows.Forms.Button
    $btnReturn.Text = "Return to Main Menu"
    $btnReturn.Location = New-Object System.Drawing.Point(50, 270)
    $btnReturn.Size = New-Object System.Drawing.Size(400, 40)
    $btnReturn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnReturn.BackColor = [System.Drawing.Color]::FromArgb(180, 0, 0) # Màu đỏ đậm
    $btnReturn.ForeColor = [System.Drawing.Color]::White
    $btnReturn.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $btnReturn.Add_Click({
        $featuresForm.Close()
    })
    $featuresForm.Controls.Add($btnReturn)

    # Show the form
    $featuresForm.ShowDialog()
}
#DONE
$buttonRenameDevice = New-GreenButton -text "Rename Device" -x 430 -y 180 -width 380 -height 60 -clickAction {
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
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $cancelButton.ForeColor = [System.Drawing.Color]::White
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $cancelButton.Size = New-Object System.Drawing.Size(200, 40)
    $cancelButton.Location = New-Object System.Drawing.Point(250, 180)
    $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelButton.Add_Click({
        $renameForm.Close()
    })
    $renameForm.Controls.Add($cancelButton)

    # Show the form
    $renameForm.ShowDialog()
}
#DONE
$buttonSetPassword = New-GreenButton -text "Set Password" -x 430 -y 260 -width 380 -height 60 -clickAction {
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
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $cancelButton.ForeColor = [System.Drawing.Color]::White
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $cancelButton.Size = New-Object System.Drawing.Size(200, 40)
    $cancelButton.Location = New-Object System.Drawing.Point(250, 230)
    $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelButton.Add_Click({
        $passwordForm.Close()
    })
    $passwordForm.Controls.Add($cancelButton)

    # Show the form
    $passwordForm.ShowDialog()
}
#DONE
$buttonJoinDomain = New-GreenButton -text "Join Domain" -x 430 -y 340 -width 380 -height 60 -clickAction {
    # Create domain/workgroup join form
    $joinForm = New-Object System.Windows.Forms.Form
    $joinForm.Text = "Join Domain/Workgroup"
    $joinForm.Size = New-Object System.Drawing.Size(500, 400)
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
    $radioDomain.Checked = $false

    $radioWorkgroup = New-Object System.Windows.Forms.RadioButton
    $radioWorkgroup.Text = "Join Workgroup"
    $radioWorkgroup.Font = New-Object System.Drawing.Font("Arial", 10)
    $radioWorkgroup.ForeColor = [System.Drawing.Color]::White
    $radioWorkgroup.Location = New-Object System.Drawing.Point(240, 30)
    $radioWorkgroup.Size = New-Object System.Drawing.Size(200, 30)
    $radioWorkgroup.BackColor = [System.Drawing.Color]::Black
    $radioWorkgroup.Checked = $true

    $groupBox.Controls.Add($radioDomain)
    $groupBox.Controls.Add($radioWorkgroup)
    $joinForm.Controls.Add($groupBox)

    # Name label
    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = "Workgroup Name:"
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
    $nameTextBox.Text = "WORKGROUP"
    $joinForm.Controls.Add($nameTextBox)

    # Username label (for domain)
    $usernameLabel = New-Object System.Windows.Forms.Label
    $usernameLabel.Text = "Username:"
    $usernameLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $usernameLabel.ForeColor = [System.Drawing.Color]::White
    $usernameLabel.Size = New-Object System.Drawing.Size(150, 30)
    $usernameLabel.Location = New-Object System.Drawing.Point(20, 270)
    $usernameLabel.Visible = $false
    $joinForm.Controls.Add($usernameLabel)

    # Username textbox
    $usernameTextBox = New-Object System.Windows.Forms.TextBox
    $usernameTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $usernameTextBox.Size = New-Object System.Drawing.Size(300, 30)
    $usernameTextBox.Location = New-Object System.Drawing.Point(170, 270)
    $usernameTextBox.BackColor = [System.Drawing.Color]::Black
    $usernameTextBox.ForeColor = [System.Drawing.Color]::Green
    $usernameTextBox.Visible = $false
    $joinForm.Controls.Add($usernameTextBox)

    # Password label (for domain)
    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Text = "Password:"
    $passwordLabel.Font = New-Object System.Drawing.Font("Arial", 12)
    $passwordLabel.ForeColor = [System.Drawing.Color]::White
    $passwordLabel.Size = New-Object System.Drawing.Size(150, 30)
    $passwordLabel.Location = New-Object System.Drawing.Point(20, 310)
    $passwordLabel.Visible = $false
    $joinForm.Controls.Add($passwordLabel)

    # Password textbox
    $passwordTextBox = New-Object System.Windows.Forms.TextBox
    $passwordTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $passwordTextBox.Size = New-Object System.Drawing.Size(300, 30)
    $passwordTextBox.Location = New-Object System.Drawing.Point(170, 310)
    $passwordTextBox.BackColor = [System.Drawing.Color]::Black
    $passwordTextBox.ForeColor = [System.Drawing.Color]::Green
    $passwordTextBox.UseSystemPasswordChar = $true
    $passwordTextBox.Visible = $false
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
    $joinButton = New-Object System.Windows.Forms.Button
    $joinButton.Text = "Join"
    $joinButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $joinButton.ForeColor = [System.Drawing.Color]::White
    $joinButton.BackColor = [System.Drawing.Color]::FromArgb(0, 180, 0)
    $joinButton.Size = New-Object System.Drawing.Size(200, 40)
    $joinButton.Location = New-Object System.Drawing.Point(30, 280)
    $joinButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
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
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $cancelButton.ForeColor = [System.Drawing.Color]::White
    $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $cancelButton.Size = New-Object System.Drawing.Size(200, 40)
    $cancelButton.Location = New-Object System.Drawing.Point(250, 280)
    $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cancelButton.Add_Click({
        $joinForm.Close()
    })
    $joinForm.Controls.Add($cancelButton)

    # Show the form
    $joinForm.ShowDialog()
}

$buttonExit = New-GreenButton -text "Exit" -x 430 -y 420 -width 380 -height 60 -clickAction {
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
