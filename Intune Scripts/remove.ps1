# Define class objects for shortcuts [do not edit]
class Shortcut {
    [system.string]$Name
}

# Define the shortcuts
$shortcuts = @(
    [Shortcut]@{Name = "Name of the shortcut to remove"}
)

# Define the icon file name
$IconFileName = "Name of the icon file to remove"

# Set default parameters for the paths and event log details
$Desktop = [Environment]::GetFolderPath("Desktop")  # Get the path to the Desktop
$AppData = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\"  # Get the path to the Start Menu Programs directory
$ProgramDataStartMenu = "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\" # Get the path to the system's Start Menu Programs directory
$IntuneProgramDirUser = "$env:APPDATA\Intune"  # Get the path to the Intune directory
$IntuneProgramDirSystem = "$env:PROGRAMDATA\Intune"  # Get the path to the Intune directory

# Loop through each shortcut in the array
foreach ($shortcut in $shortcuts) {
    # Define the path to the desktop shortcut
    $desktopShortcutFile = "$Desktop\$($shortcut.Name).lnk"
    if (Test-Path $desktopShortcutFile) {
        # Remove the desktop shortcut
        Remove-Item $desktopShortcutFile
    }

    # Define the path to the User Start Menu shortcut
    $userStartMenuShortcutFile = "$AppData\$($shortcut.Name).lnk"
    if (Test-Path $userStartMenuShortcutFile) {
        # Remove the User Start Menu shortcut
        Remove-Item $userStartMenuShortcutFile
    }

    # Define the path to the System Start Menu shortcut
    $systemStartMenuShortcutFile = "$ProgramDataStartMenu\$($shortcut.Name).lnk"
    if (Test-Path $systemStartMenuShortcutFile) {
        # Remove the System Start Menu shortcut
        Remove-Item $systemStartMenuShortcutFile
    }
}

# Define the path to the shortcut icon file
$ShortcutIconFileUser = "$IntuneProgramDirUser\$IconFileName.ico"
if (Test-Path $ShortcutIconFileUser) {
    # Remove the shortcut icon file
    Remove-Item $ShortcutIconFileUser
}

# Define the path to the shortcut icon file
$ShortcutIconFileSystem = "$IntuneProgramDirSystem\$IconFileName.ico"
if (Test-Path $ShortcutIconFileSystem) {
    # Remove the shortcut icon file
    Remove-Item $ShortcutIconFileSystem
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
    $desktopShortcutFile = "$Desktop\$($shortcut.Name).lnk"
    $userStartMenuShortcutFile = "$AppData\$($shortcut.Name).lnk"
    $systemStartMenuShortcutFile = "$ProgramDataStartMenu\$($shortcut.Name).lnk"
    if ((Test-Path $desktopShortcutFile) -or (Test-Path $userStartMenuShortcutFile) -or (Test-Path $systemStartMenuShortcutFile)) {
        $message = "Failed to remove Shortcut $($shortcut.Name).lnk"
        $entryType = "Error"
    } else {
        $message = "Shortcut $($shortcut.Name).lnk successfully Removed"
        $entryType = "Information"
    }

    # Write the entry to the event log
    Write-EventLog -LogName $logName -Source $source -EventId $eventID -EntryType $entryType -Message $message
}