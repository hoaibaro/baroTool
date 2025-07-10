# ADMIN PRIVILEGES CHECK & INITIALIZATION
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "This script requires administrative privileges. Attempting to restart with elevation..."

    # Restart script with admin privileges
    $scriptPath = $MyInvocation.MyCommand.Path
    $arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""

    Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs

    # Exit the current non-elevated instance
    exit
}

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
# Hide Main Menu
function Hide-MainMenu {
    $script:form.Hide()
}

# Show Main Menu
function Show-MainMenu {
    $script:form.Show()
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

# THÊM vào trước phần tạo form
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop

# CREATE MAIN FORM - RESIZABLE VERSION
$script:form = New-Object System.Windows.Forms.Form
$script:form.Text = "BAOPROVIP - SYSTEM MANAGEMENT"
$script:form.Size = New-Object System.Drawing.Size(850, 5)
$script:form.MinimumSize = New-Object System.Drawing.Size(800, 550)  # Kích thước tối thiểu
$script:form.StartPosition = "CenterScreen"
$script:form.BackColor = [System.Drawing.Color]::Black
$script:form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable  # CHỖ NÀY THAY ĐỔI
$script:form.MaximizeBox = $true  # Cho phép maximize

# Add gradient background với resize handling
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

# Title - RESPONSIVE
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

function Show-SetPasswordForm {
        param(
            [string]$currentUser
        )
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Set Password"
        $form.Size = New-Object System.Drawing.Size(500, 270)
        $form.StartPosition = "CenterScreen"
        $form.BackColor = [System.Drawing.Color]::Black
        $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false
    
        # Title
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "Set Password for Current User"
        $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 255, 0)
        $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $titleLabel.Size = New-Object System.Drawing.Size(480, 40)
        $titleLabel.Location = New-Object System.Drawing.Point(10, 20)
        $form.Controls.Add($titleLabel)
    
        # User label
        $userLabel = New-Object System.Windows.Forms.Label
        $userLabel.Text = "Current User: $currentUser"
        $userLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
        $userLabel.ForeColor = [System.Drawing.Color]::White
        $userLabel.Size = New-Object System.Drawing.Size(480, 30)
        $userLabel.Location = New-Object System.Drawing.Point(20, 70)
        $form.Controls.Add($userLabel)
    
        # Password label
        $passwordLabel = New-Object System.Windows.Forms.Label
        $passwordLabel.Text = "New Password:"
        $passwordLabel.Font = New-Object System.Drawing.Font("Arial", 12)
        $passwordLabel.ForeColor = [System.Drawing.Color]::White
        $passwordLabel.Size = New-Object System.Drawing.Size(150, 30)
        $passwordLabel.Location = New-Object System.Drawing.Point(20, 110)
        $form.Controls.Add($passwordLabel)
    
        # Password textbox
        $passwordTextBox = New-Object System.Windows.Forms.TextBox
        $passwordTextBox.Font = New-Object System.Drawing.Font("Arial", 12)
        $passwordTextBox.Size = New-Object System.Drawing.Size(220, 30)
        $passwordTextBox.Location = New-Object System.Drawing.Point(170, 110)
        $passwordTextBox.BackColor = [System.Drawing.Color]::White
        $passwordTextBox.ForeColor = [System.Drawing.Color]::Black
        $passwordTextBox.UseSystemPasswordChar = $false # Mặc định hiển thị password
        $form.Controls.Add($passwordTextBox)
    
        # Show Password checkbox (default checked)
        $showPasswordCheckBox = New-Object System.Windows.Forms.CheckBox
        $showPasswordCheckBox.Text = "Show"
        $showPasswordCheckBox.Location = New-Object System.Drawing.Point(400, 115)
        $showPasswordCheckBox.Size = New-Object System.Drawing.Size(100, 20)
        $showPasswordCheckBox.ForeColor = [System.Drawing.Color]::White
        $showPasswordCheckBox.Font = New-Object System.Drawing.Font("Arial", 9)
        $showPasswordCheckBox.BackColor = [System.Drawing.Color]::Transparent
        $showPasswordCheckBox.Checked = $true
        $showPasswordCheckBox.Add_CheckedChanged({
            $passwordTextBox.UseSystemPasswordChar = -not $showPasswordCheckBox.Checked
        })
        $form.Controls.Add($showPasswordCheckBox)
    
        # Info label for empty password
        $infoLabel = New-Object System.Windows.Forms.Label
        $infoLabel.Text = "Leave the password field empty to set a blank password."
        $infoLabel.Font = New-Object System.Drawing.Font("Arial", 9)
        $infoLabel.ForeColor = [System.Drawing.Color]::Silver
        $infoLabel.Size = New-Object System.Drawing.Size(450, 20)
        $infoLabel.Location = New-Object System.Drawing.Point(20, 145)
        $form.Controls.Add($infoLabel)
    
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
                    $command = "net user $currentUser """""
                } else {
                    $command = "net user $currentUser $password"
                }
                $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command" -NoNewWindow -Wait -PassThru
                if ($process.ExitCode -eq 0) {
                    if ([string]::IsNullOrEmpty($password)) {
                        [System.Windows.Forms.MessageBox]::Show("Password has been removed. User '$currentUser' can now log in without a password.", "Password Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Password has been changed.", "Password Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    }
                    $form.Close()
                } else {
                    throw "Failed to set password. Exit code: $($process.ExitCode)"
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error setting password: $_`n`nNote: This operation requires administrative privileges.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })
        $form.Controls.Add($setButton)
    
        # Cancel button
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
        $cancelButton.ForeColor = [System.Drawing.Color]::White
        $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(150, 0, 0)
        $cancelButton.Size = New-Object System.Drawing.Size(200, 40)
        $cancelButton.Location = New-Object System.Drawing.Point(250, 180)
        $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $cancelButton.Add_Click({
            $form.Close()
        })
        $form.Controls.Add($cancelButton)
    
        # Set Accept/Cancel button for Enter/Esc
        $form.AcceptButton = $setButton
        $form.CancelButton = $cancelButton
    
        # Focus on password textbox when form shows
        $form.Add_Shown({
            $passwordTextBox.Focus()
        })
    
        # Show the form
        $form.ShowDialog()
} 

function Set-UserPassword {
    param(
        [string]$user,
        [string]$password
    )
    try {
        if ([string]::IsNullOrEmpty($password)) {
            # Xóa mật khẩu (blank)
            $command = "net user $user """""
        } else {
            $command = "net user $user $password"
        }
        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command" -NoNewWindow -Wait -PassThru
        return $process.ExitCode -eq 0
    } catch {
        return $false
    }
}

function Remove-UserPassword {
    param(
        [string]$user
    )
    try {
        $command = "net user $user """""
        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command" -NoNewWindow -Wait -PassThru
        return $process.ExitCode -eq 0
    } catch {
        return $false
    }
}

function Invoke-SetPasswordDialog {
    $currentUser = $env:USERNAME
    Hide-MainMenu
    $result = Show-SetPasswordForm -currentUser $currentUser

    if ($result.Action -eq "set") {
        $success = Set-UserPassword -user $currentUser -password $result.Password
        if ($success) {
            if ([string]::IsNullOrEmpty($result.Password)) {
                [System.Windows.Forms.MessageBox]::Show("Password has been removed. User '$currentUser' can now log in without a password.", "Password Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                [System.Windows.Forms.MessageBox]::Show("Password has been changed.", "Password Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Error setting password. This operation may require administrative privileges.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } elseif ($result.Action -eq "remove") {
        $success = Remove-UserPassword -user $currentUser
        if ($success) {
            [System.Windows.Forms.MessageBox]::Show("Password has been removed. User '$currentUser' can now log in without a password.", "Password Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Error removing password. This operation may require administrative privileges.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    Show-MainMenu
}

$buttonSetPassword = New-DynamicButton -text "[8] Set Password" -x 430 -y 260 -width 380 -height 60 -clickAction {
    Invoke-SetPasswordDialog
}

$script:form.Controls.Add($buttonSetPassword)

# Start Application
$script:form.ShowDialog() 