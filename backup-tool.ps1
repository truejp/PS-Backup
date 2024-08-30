# Import the required .NET types
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$boxBackground = "#cccfce" # Box Background Color
$buttonBackColor = "#525753" # Button Background Color
$buttonTextColor = "#e9f2eb" # Button Text Color

# Autostart Configuration
# Bestimme den Pfad zum aktuellen Skript
$scriptPath = $MyInvocation.MyCommand.Path

# Bestimme den Pfad zur Batch-Datei im Autostart-Ordner
$autostartFolder = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\Start Menu\Programs\Startup")
$batchFilePath = [System.IO.Path]::Combine($autostartFolder, "Run_PS_Backup.bat")

# Funktion zum Überprüfen, ob die Batch-Datei existiert
function Test-BatchFileExists {
    return Test-Path $batchFilePath
}

# Funktion zum Erstellen der Batch-Datei
function Create-BatchFile {
    $batchContent = "@echo off`r`npowershell.exe -ExecutionPolicy Bypass -File `"$scriptPath`""
    Set-Content -Path $batchFilePath -Value $batchContent
}

# Überprüfen, ob die Batch-Datei existiert und sie erstellen, falls nicht
if (-not (Test-BatchFileExists)) {
    Create-BatchFile
    Write-Host "Batch-Datei erstellt und im Autostart-Ordner platziert."
} else {
    Write-Host "Batch-Datei existiert bereits im Autostart-Ordner."
}



# Function to save settings
function Save-Settings {
    param (
        [string]$filePath,
        [pscustomobject]$settings
    )
    $settings | ConvertTo-Json -Depth 10 | Set-Content -Path $filePath -Encoding utf8
    Log-Event -message "Settings saved."
}

# Function to load settings
function Load-Settings {
    param (
        [string]$filePath
    )
    if (Test-Path $filePath) {
        $settings = Get-Content -Path $filePath | ConvertFrom-Json
    } else {
        # Default values if no file exists
        $settings = [pscustomobject]@{
            BackupPaths           = @()
            BackupFrequency       = "Daily"
            IsActive              = $false
            LastBackup            = (Get-Date).ToString("o")  # Initialize with current date
            EnableNotifications   = $true  # Default to enabled
        }
    }

    # Ensure LastBackup exists
    if (-not $settings.PSObject.Properties.Match('LastBackup')) {
        $settings | Add-Member -MemberType NoteProperty -Name LastBackup -Value (Get-Date).ToString("o")
    }

    return $settings
}


function Start-Backup {
    Log-Event -message "Backup plan started with frequency $($settings.BackupFrequency)."
    Show-Notification -title "Backup Plan Started" -message "Backup plan started with the $($settings.BackupFrequency) preset."
}



# Function to update status
function Update-Status {
    if ($settings.IsActive) {
        try {
            $lastBackupDate = [datetime]::Parse($settings.LastBackup)
        } catch {
            $lastBackupDate = Get-Date
            $settings.LastBackup = $lastBackupDate.ToString("o")  # If parse fails, use current time
        }

        $nextBackup = if ($settings.BackupFrequency -eq 'Hourly') {
            $lastBackupDate.AddHours(1)
        } else {
            $lastBackupDate.AddDays(1)
        }

        $currentTime = Get-Date
        $timeRemaining = $nextBackup - $currentTime

        # Zeit in Stunden und Minuten umrechnen
        $hoursRemaining = [math]::Floor($timeRemaining.TotalHours)
        $minutesRemaining = $timeRemaining.Minutes

        if ($timeRemaining -lt [timespan]::Zero) {
            # Wenn die Zeit seit dem letzten Backup größer ist als die Frequenz, führe ein Backup durch
            $statusLabel.Text = "Protection Plan Active"
            $nextBackupLabel.Text = "Planned: $($nextBackup.ToString('g')) ($hoursRemaining h $minutesRemaining min late)"
            Run-Backup
        } else {
            $statusLabel.Text = "Protection Plan Active"
            $nextBackupLabel.Text = "Planned: $($nextBackup.ToString('g')) ($hoursRemaining h $minutesRemaining min remaining)"
        }
    } else {
        $statusLabel.Text = "Protection Plan Disabled"
        $nextBackupLabel.Text = "Your data is not being backed up."
    }
}


# Function to perform the backup
function Run-Backup {
    Log-Event -message "Backup started."
    Show-Notification -title "Backup Started" -message "Backup started."
    $statusLabel.Text = "Running backup... Please wait."
    $settings.BackupPaths | ForEach-Object {
        $source = $_.Source
        $destination = $_.Destination
        Log-Event -message "Starting backup from $source to $destination."
        Show-Notification -title "Sub-Task in Progress" -message "Starting backup from $source to $destination."
        robocopy $source $destination /MIR /Z /R:5 /W:15
        Log-Event -message "Backup from $source to $destination completed."
        Show-Notification -title "Sub-Task Completed" -message "Backup from $source to $destination completed."
    }
    Show-Notification -title "Backup Completed" -message "Your data has been successfully backed up."
    $settings.LastBackup = (Get-Date).ToString("o")  # ISO 8601 Format
    Save-Settings -filePath $settingsFile -settings $settings
    $statusLabel.Text = "Backup completed."
    Update-Status
}

# Function to log events
function Log-Event {
    param (
        [string]$message
    )
    $logDir = Join-Path $PSScriptRoot "logs"
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory | Out-Null
    }

    $logFilePath = Join-Path $logDir "$(Get-Date -Format 'yyyy-MM-dd').log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Add-Content -Path $logFilePath -Value $logMessage
}

# Function to show notifications
function Show-Notification {
    param (
        [string]$title,
        [string]$message
    )
    if ($settings.EnableNotifications) {
        $notifyIcon.BalloonTipTitle = $title
        $notifyIcon.BalloonTipText = $message
        $notifyIcon.ShowBalloonTip(3000)  # Time in milliseconds
    }
}


# Function to open log folder
function Open-LogFolder {
    $logDir = Join-Path $PSScriptRoot "logs"
    if (Test-Path $logDir) {
        Start-Process explorer.exe $logDir
    } else {
        [System.Windows.Forms.MessageBox]::Show("Log folder does not exist.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to browse for a folder
function Browse-Folder {
    param (
        [string]$title
    )
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $title
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    return $null
}

# Timer für Status-Aktualisierung
$statusUpdateTimer = New-Object System.Windows.Forms.Timer
$statusUpdateTimer.Interval = 60000 # Aktualisiert alle 60 Sekunden
$statusUpdateTimer.Add_Tick({
    Update-Status
})
$statusUpdateTimer.Start()

# Create GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = "PS Backup"  # Changed title
$form.Size = New-Object System.Drawing.Size(330,230)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::White  # Background color

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10,10)
$statusLabel.Size = New-Object System.Drawing.Size(330,20)
$statusLabel.Font = New-Object System.Drawing.Font('Arial', 10)
$statusLabel.TextAlign = 'MiddleLeft'
$statusLabel.ForeColor = [System.Drawing.Color]::DarkBlue  # Text color

# Next Backup Label
$nextBackupLabel = New-Object System.Windows.Forms.Label
$nextBackupLabel.Location = New-Object System.Drawing.Point(10,35)
$nextBackupLabel.Size = New-Object System.Drawing.Size(330,20)
$nextBackupLabel.Font = New-Object System.Drawing.Font('Arial', 8)
$nextBackupLabel.TextAlign = 'MiddleLeft'
$nextBackupLabel.ForeColor = [System.Drawing.Color]::DarkBlue  # Text color

# NotifyIcon for notifications
$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
$notifyIcon.Visible = $true

# Version and Author Label
$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Location = New-Object System.Drawing.Point(10,170)
$versionLabel.Size = New-Object System.Drawing.Size(330,20)
$versionLabel.Font = New-Object System.Drawing.Font('Arial', 8)
$versionLabel.TextAlign = 'MiddleLeft'
$versionLabel.Text = "PSB v1.1@2024 | Author: Philipp Lehnet"
$versionLabel.ForeColor = [System.Drawing.Color]::DarkGray  # Text color

# Load settings
$settingsFile = "$PSScriptRoot\backup_settings.json"
$settings = Load-Settings -filePath $settingsFile

# Backup Timer (initially not active)
if ($settings.IsActive) {
    Start-Backup
}

# Update status
Update-Status

# Activate Button
$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "Activate"
$startButton.Size = New-Object System.Drawing.Size(120,30)
$startButton.Location = New-Object System.Drawing.Point(10,70)
$startButton.BackColor = [System.Drawing.Color]::LightGreen  # Button color
$startButton.ForeColor = [System.Drawing.Color]::Black # Text color
$startButton.Add_Click({
    if (-not $settings.IsActive) {
        Start-Backup
        $settings.IsActive = $true
        Save-Settings -filePath $settingsFile -settings $settings
        Update-Status
    }
})

# Deactivate Button
$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Text = "Deactivate"
$stopButton.Size = New-Object System.Drawing.Size(120,30)
$stopButton.Location = New-Object System.Drawing.Point(140,70)
$stopButton.BackColor = [System.Drawing.Color]::LightCoral  # Button color
$stopButton.ForeColor = [System.Drawing.Color]::Black # Text color
$stopButton.Add_Click({
    if ($settings.IsActive) {
        $settings.IsActive = $false
        Save-Settings -filePath $settingsFile -settings $settings
        Update-Status
    }
})

# Run Now Button
$runNowButton = New-Object System.Windows.Forms.Button
$runNowButton.Text = "Run Now"
$runNowButton.Size = New-Object System.Drawing.Size(120,30)
$runNowButton.Location = New-Object System.Drawing.Point(140,110)
$runNowButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonBackColor")  # Button color
$runNowButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonTextColor")  # Text color
$runNowButton.Add_Click({
    Run-Backup
})

# Settings Button
$settingsButton = New-Object System.Windows.Forms.Button
$settingsButton.Text = "Settings"
$settingsButton.Size = New-Object System.Drawing.Size(120,30)
$settingsButton.Location = New-Object System.Drawing.Point(10,110)
$settingsButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonBackColor")  # Button color
$settingsButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonTextColor")  # Text color
$settingsButton.Add_Click({
    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = "Settings"
    $settingsForm.Size = New-Object System.Drawing.Size(550,700)
    $settingsForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $settingsForm.MaximizeBox = $false
    $settingsForm.BackColor = [System.Drawing.Color]::LightGray  # Background color

    # GroupBox for Backup Paths
    $backupPathsGroupBox = New-Object System.Windows.Forms.GroupBox
    $backupPathsGroupBox.Text = "Backup Paths"
    $backupPathsGroupBox.Location = New-Object System.Drawing.Point(10,10)
    $backupPathsGroupBox.Size = New-Object System.Drawing.Size(520,160)
    $backupPathsGroupBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$boxBackground")  # Background color

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10,20)
    $listBox.Size = New-Object System.Drawing.Size(500,100)
    $settings.BackupPaths | ForEach-Object {
        $listBox.Items.Add("Source: $($_.Source) -> Destination: $($_.Destination)")
    }

    # GroupBox for Source Directory
    $sourceGroupBox = New-Object System.Windows.Forms.GroupBox
    $sourceGroupBox.Text = "Set Source Directory"
    $sourceGroupBox.Location = New-Object System.Drawing.Point(10,180)
    $sourceGroupBox.Size = New-Object System.Drawing.Size(520,120)
    $sourceGroupBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$boxBackground")  # Background color

    $sourceLabel = New-Object System.Windows.Forms.Label
    $sourceLabel.Text = "Source Path"
    $sourceLabel.Location = New-Object System.Drawing.Point(10,20)
    $sourceLabel.Size = New-Object System.Drawing.Size(200,20)
    $sourceLabel.ForeColor = [System.Drawing.Color]::DarkBlue  # Text color

    $sourceTextBox = New-Object System.Windows.Forms.TextBox
    $sourceTextBox.Location = New-Object System.Drawing.Point(10,45)
    $sourceTextBox.Size = New-Object System.Drawing.Size(320,20)

    $browseSourceButton = New-Object System.Windows.Forms.Button
    $browseSourceButton.Text = "Browse"
    $browseSourceButton.Location = New-Object System.Drawing.Point(340,45)
    $browseSourceButton.Size = New-Object System.Drawing.Size(100,30)
    $browseSourceButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonBackColor")  # Button color
    $browseSourceButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonTextColor")  # Text color
    $browseSourceButton.Add_Click({
        $sourcePath = Browse-Folder -title "Select Source Folder"
        if ($sourcePath) {
            $sourceTextBox.Text = $sourcePath
        }
    })

    # GroupBox for Destination Directory
    $destinationGroupBox = New-Object System.Windows.Forms.GroupBox
    $destinationGroupBox.Text = "Set Destination Directory"
    $destinationGroupBox.Location = New-Object System.Drawing.Point(10,310)
    $destinationGroupBox.Size = New-Object System.Drawing.Size(520,160)
    $destinationGroupBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$boxBackground")  # Background color

    $destinationLabel = New-Object System.Windows.Forms.Label
    $destinationLabel.Text = "Destination Path"
    $destinationLabel.Location = New-Object System.Drawing.Point(10,20)
    $destinationLabel.Size = New-Object System.Drawing.Size(200,20)
    $destinationLabel.ForeColor = [System.Drawing.Color]::DarkBlue  # Text color

    $destinationTextBox = New-Object System.Windows.Forms.TextBox
    $destinationTextBox.Location = New-Object System.Drawing.Point(10,45)
    $destinationTextBox.Size = New-Object System.Drawing.Size(320,20)

    $browseDestinationButton = New-Object System.Windows.Forms.Button
    $browseDestinationButton.Text = "Browse"
    $browseDestinationButton.Location = New-Object System.Drawing.Point(340,45)
    $browseDestinationButton.Size = New-Object System.Drawing.Size(100,30)
    $browseDestinationButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonBackColor")  # Button color
    $browseDestinationButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonTextColor")  # Text color
    $browseDestinationButton.Add_Click({
        $destinationPath = Browse-Folder -title "Select Destination Folder"
        if ($destinationPath) {
            $destinationTextBox.Text = $destinationPath
        }
    })

    # Add and Remove Buttons (Moved to Destination GroupBox)
    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Text = "Add"
    $addButton.Location = New-Object System.Drawing.Point(10,80)
    $addButton.Size = New-Object System.Drawing.Size(100,30)
    $addButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonBackColor")  # Button color
    $addButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonTextColor")  # Text color
    $addButton.Add_Click({
        $source = $sourceTextBox.Text
        $destination = $destinationTextBox.Text
        if ($source -and $destination) {
            # Ensure BackupPaths is an array
            if (-not $settings.BackupPaths) {
                $settings.BackupPaths = @()
            }
            # Add new backup path
            $settings.BackupPaths += [pscustomobject]@{ Source = $source; Destination = $destination }
            # Save settings to file
            Save-Settings -filePath $settingsFile -settings $settings
            $listBox.Items.Add("Source: $source -> Destination: $destination")
            $sourceTextBox.Clear()
            $destinationTextBox.Clear()
            Log-Event -message "Backup path added: Source = $source, Destination = $destination."
        }
    })

    $removeButton = New-Object System.Windows.Forms.Button
    $removeButton.Text = "Remove"
    $removeButton.Location = New-Object System.Drawing.Point(120,80)
    $removeButton.Size = New-Object System.Drawing.Size(100,30)
    $removeButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonBackColor")  # Button color
    $removeButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonTextColor")  # Text color
    $removeButton.Add_Click({
        if ($listBox.SelectedIndex -ne -1) {
            $item = $listBox.SelectedItem
            $source, $destination = $item -replace 'Source: (.+) -> Destination: (.+)', '$1,$2' -split ','
            $settings.BackupPaths = $settings.BackupPaths | Where-Object { $_.Source -ne $source -and $_.Destination -ne $destination }
            Save-Settings -filePath $settingsFile -settings $settings
            $listBox.Items.RemoveAt($listBox.SelectedIndex)
            Log-Event -message "Backup path removed: Source = $source, Destination = $destination."
        }
    })

    # GroupBox for Backup Frequency
    $frequencyGroupBox = New-Object System.Windows.Forms.GroupBox
    $frequencyGroupBox.Text = "Backup Settings"
    $frequencyGroupBox.Location = New-Object System.Drawing.Point(10,480)
    $frequencyGroupBox.Size = New-Object System.Drawing.Size(520,120)
    $frequencyGroupBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$boxBackground")  # Background color

    $frequencyLabel = New-Object System.Windows.Forms.Label
    $frequencyLabel.Text = "Run Backups"
    $frequencyLabel.Location = New-Object System.Drawing.Point(10,20)
    $frequencyLabel.Size = New-Object System.Drawing.Size(120,20)
    $frequencyLabel.ForeColor = [System.Drawing.Color]::DarkBlue  # Text color

    $frequencyComboBox = New-Object System.Windows.Forms.ComboBox
    $frequencyComboBox.Location = New-Object System.Drawing.Point(10,40)
    $frequencyComboBox.Size = New-Object System.Drawing.Size(150,20)
    $frequencyComboBox.Items.AddRange(@("Hourly", "Daily"))
    $frequencyComboBox.SelectedItem = $settings.BackupFrequency

    $saveFrequencyButton = New-Object System.Windows.Forms.Button
    $saveFrequencyButton.Text = "Save Frequency"
    $saveFrequencyButton.Location = New-Object System.Drawing.Point(170,35)
    $saveFrequencyButton.Size = New-Object System.Drawing.Size(120,30)
    $saveFrequencyButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonBackColor")  # Button color
    $saveFrequencyButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonTextColor")  # Text color
    $saveFrequencyButton.Add_Click({
        $settings.BackupFrequency = $frequencyComboBox.SelectedItem
        Save-Settings -filePath $settingsFile -settings $settings
        Update-Status
        Log-Event -message "Backup frequency changed to $($settings.BackupFrequency)."
    })

    $notificationsCheckBox = New-Object System.Windows.Forms.CheckBox
    $notificationsCheckBox.Text = "Enable Notifications"
    $notificationsCheckBox.Location = New-Object System.Drawing.Point(10,80)
    $notificationsCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $notificationsCheckBox.Checked = $settings.EnableNotifications

    $saveNotificationsButton = New-Object System.Windows.Forms.Button
    $saveNotificationsButton.Text = "Save Notifications"
    $saveNotificationsButton.Location = New-Object System.Drawing.Point(170,75)
    $saveNotificationsButton.Size = New-Object System.Drawing.Size(120,30)
    $saveNotificationsButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonBackColor")  # Button color
    $saveNotificationsButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonTextColor")  # Text color
    $saveNotificationsButton.Add_Click({
        $settings.EnableNotifications = $notificationsCheckBox.Checked
        Save-Settings -filePath $settingsFile -settings $settings
        Log-Event -message "Notification settings changed to $($settings.EnableNotifications)."
    })

    # Open Log Folder Button
    $openLogFolderButton = New-Object System.Windows.Forms.Button
    $openLogFolderButton.Text = "Open Log Folder"
    $openLogFolderButton.Location = New-Object System.Drawing.Point(10,610)
    $openLogFolderButton.Size = New-Object System.Drawing.Size(200,30)
    $openLogFolderButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonBackColor")  # Button color
    $openLogFolderButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("$buttonTextColor")  # Text color
    $openLogFolderButton.Add_Click({
        Open-LogFolder
    })


    # Add controls to settings form
    $settingsForm.Controls.Add($backupPathsGroupBox)
    $backupPathsGroupBox.Controls.Add($listBox)

    $settingsForm.Controls.Add($sourceGroupBox)
    $sourceGroupBox.Controls.Add($sourceLabel)
    $sourceGroupBox.Controls.Add($sourceTextBox)
    $sourceGroupBox.Controls.Add($browseSourceButton)

    $settingsForm.Controls.Add($destinationGroupBox)
    $destinationGroupBox.Controls.Add($destinationLabel)
    $destinationGroupBox.Controls.Add($destinationTextBox)
    $destinationGroupBox.Controls.Add($browseDestinationButton)
    $destinationGroupBox.Controls.Add($addButton)
    $destinationGroupBox.Controls.Add($removeButton)

    $settingsForm.Controls.Add($frequencyGroupBox)
    $frequencyGroupBox.Controls.Add($frequencyLabel)
    $frequencyGroupBox.Controls.Add($frequencyComboBox)
    $frequencyGroupBox.Controls.Add($saveFrequencyButton)
    $frequencyGroupBox.Controls.Add($notificationsCheckBox)
    $frequencyGroupBox.Controls.Add($saveNotificationsButton)

    $settingsForm.Controls.Add($openLogFolderButton)
    
    $settingsForm.ShowDialog()
})

$form.Controls.Add($statusLabel)
$form.Controls.Add($nextBackupLabel)
$form.Controls.Add($startButton)
$form.Controls.Add($stopButton)
$form.Controls.Add($runNowButton)
$form.Controls.Add($settingsButton)
$form.Controls.Add($versionLabel)

# Set form background color
$form.BackColor = [System.Drawing.Color]::White  # Background color

$form.Text = "PSB - PowerShell Backup"  # Title update
$form.ShowDialog()
