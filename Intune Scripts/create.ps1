# Define class objects for shortcuts [do not edit]
class Shortcut {
    [system.string]$Name
    [system.string]$TargetPath
    [system.string]$Arguments = ""
    [system.string]$WorkingDirectory
}

# Determines if a shortcut should be created on the user's desktop [Win32app deployment should be targeting user] [cannot be mixed with system context]
$ShortcutOnDesktop          = $False
# Determines if a shortcut should be created in the user's start menu [Win32app deployment should be targeting user] [cannot be mixed with system context]
$ShortcutInUserStartMenu    = $False
# Determines if a shortcut should be created in the system start menu [Win32app deployment should be targeting system] [cannot be mixed with user context]
$ShortcutInSystemStartMenu  = $True

# Define the shortcuts
$shortcuts = @(
    [Shortcut]@{
        Name = "Name of the shortcut to create"
        TargetPath = "Path to the target executable"
        Arguments = "Arguments to pass to the target executable, Remove this line if no arguments are needed"
        WorkingDirectory = "Path to the working directory of the target executable"
    }
)

# Define the name of the icon file
$IconFileName      = "Name of the icon file to create"

# Base64 representation of the icon you can convert ico file to base64 using https://base64.guru/converter/encode/image/ico
$IconBase64 = "String of base64 encoded icon file"

# Get the path to the current user's Desktop directory
$Desktop = [Environment]::GetFolderPath("Desktop")
# Get the path to the current user's Start Menu Programs directory
$AppData = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\"
# Get the path to the system's Start Menu Programs directory
$ProgramDataStartMenu = "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\"
# Define the path to the Intune directory within the current user's AppData
$IntuneProgramDirUser = "$env:APPDATA\Intune"
# Define the path to the Intune directory within the system ProgramData
$IntuneProgramDirSystem = "$env:PROGRAMDATA\Intune"
# Define the path to the temporary icon file within the Intune directory
$TempIconUser = "$IntuneProgramDirUser\$IconFileName.ico"
# Define the path to the temporary icon file within the Intune directory
$TempIconSystem = "$IntuneProgramDirSystem\$IconFileName.ico"

# Check if the Intune directory exists
if (!(Test-Path -Path $IntuneProgramDirUser)) {
    # If not, create it
    New-Item -ItemType Directory -Path $IntuneProgramDirUser
}

# Check if the Intune directory exists
if (!(Test-Path -Path $IntuneProgramDirSystem)) {
    # If not, create it
    New-Item -ItemType Directory -Path $IntuneProgramDirSystem
}

# Check if $ShortcutInUserStartMenu is $true
if ($ShortcutInUserStartMenu) {
    # Check if the icon file exists
    if(-not (Test-Path $TempIconUser)) {
        # If not, convert the base64 string to bytes and write them to a file
        [byte[]]$Bytes = [convert]::FromBase64String($IconBase64)
        [System.IO.File]::WriteAllBytes($TempIconUser,$Bytes)
    }
}

# Check if $ShortcutInSystemStartMenu is $true
if ($ShortcutInSystemStartMenu) {
    # Check if the icon file exists
    if(-not (Test-Path $TempIconSystem)) {
        # If not, convert the base64 string to bytes and write them to a file
        [byte[]]$Bytes = [convert]::FromBase64String($IconBase64)
        [System.IO.File]::WriteAllBytes($TempIconSystem,$Bytes)
    }
}

# Iterate over each shortcut
foreach ($shortcut in $shortcuts) {
    # Test if the shortcut is currently present on the desktop
    $desktopShortcutPresent = Get-ChildItem -Path $Desktop | Where-Object {$_.Name -eq "$($shortcut.Name).lnk"}
    # Test if the shortcut is currently present in the user's start menu
    $UserStartMenuShortcutPresent = Get-ChildItem -Path $AppData | Where-Object {$_.Name -eq "$($shortcut.Name).lnk"}
    # Test if the shortcut is currently present in the system start menu
    $SystemStartMenuShortcutPresent = Get-ChildItem -Path $ProgramDataStartMenu | Where-Object {$_.Name -eq "$($shortcut.Name).lnk"}

    # If it is, delete it so we can update it
    if ($null -ne $desktopShortcutPresent) {
        Remove-Item $desktopShortcutPresent.VersionInfo.FileName -Force -Confirm:$False
    }
    if ($null -ne $UserStartMenuShortcutPresent) {
        Remove-Item $UserStartMenuShortcutPresent.VersionInfo.FileName -Force -Confirm:$False
    }
    if ($null -ne $SystemStartMenuShortcutPresent) {
        Remove-Item $SystemStartMenuShortcutPresent.VersionInfo.FileName -Force -Confirm:$False
    }
}

# Create a new COM object for creating shortcuts
$WScriptShell = New-Object -ComObject WScript.Shell

# Define the locations and corresponding icons
$locations = @()
if ($ShortcutOnDesktop) {
    $locations += @{Path = $Desktop; Icon = $TempIconUser}
}
if ($ShortCutInUserStartMenu) {
    $locations += @{Path = $AppData; Icon = $TempIconUser}
}
if ($ShortCutInSystemStartMenu) {
    $locations += @{Path = $ProgramDataStartMenu; Icon = $TempIconSystem}
}

# Iterate over each location
foreach ($location in $locations) {
    # Iterate over each shortcut
    foreach ($shortcut in $shortcuts) {
        # Define the path for the shortcut
        $shortcutFile = "$($location.Path)\$($shortcut.Name).lnk"

        # Create the shortcut
        $shortcutObject = $WScriptShell.CreateShortcut($shortcutFile)
        $shortcutObject.TargetPath = $shortcut.TargetPath
        $shortcutObject.Arguments = $shortcut.Arguments
        $shortcutObject.WorkingDirectory = $shortcut.WorkingDirectory
        $shortcutObject.IconLocation = $location.Icon

        # Save the shortcut
        $shortcutObject.Save()
    }
}

# Define the source, log name, event ID, and entry type for the event log
$source = "Intune Shortcut Script"
$logName = "Application"
$eventID = 1001
$entryType = "Information"

# Check if the event log source exists
if (![Diagnostics.EventLog]::SourceExists($source)) {
    # If not, create it
    [Diagnostics.EventLog]::CreateEventSource($source, $logName)
}

# Iterate over each shortcut
foreach ($shortcut in $shortcuts) {
    # Define the message for the event log
    $message = "Shortcut $($shortcut.Name).lnk created successfully"

    # Write the entry to the event log
    Write-EventLog -LogName $logName -Source $source -EventId $eventID -EntryType $entryType -Message $message
}