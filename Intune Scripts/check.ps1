class Shortcut {
    [system.string]$Name
}

# Define the shortcuts
$shortcuts = @(
    [Shortcut]@{Name = "Name of the shortcut to check for"}
)

# Define whether to check for the shortcuts on the desktop, user's start menu, and system's start menu
$ShortcutOnDesktop = $false
$ShortcutOnUserStartMenu = $false
$ShortcutOnSystemStartMenu = $true
# Get the path to the current user's Desktop directory
$Desktop = [Environment]::GetFolderPath("Desktop")
# Get the path to the current user's Start Menu Programs directory
$UserStartMenu = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
# Get the path to the system's Start Menu Programs directory
$SystemStartMenu = "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs"

# Initialize the list of shortcut files
$shortcutFiles = @()

# Loop through each shortcut in the array
foreach ($shortcut in $shortcuts) {
    # Add the shortcuts on the desktop to the list if $ShortcutOnDesktop is $true
    if ($ShortcutOnDesktop) {
        $shortcutFiles += "$Desktop\$($shortcut.Name).lnk"
    }

    # Add the shortcuts on the user's start menu to the list if $ShortcutOnUserStartMenu is $true
    if ($ShortcutOnUserStartMenu) {
        $shortcutFiles += "$UserStartMenu\$($shortcut.Name).lnk"
    }

    # Add the shortcuts on the system's start menu to the list if $ShortcutOnSystemStartMenu is $true
    if ($ShortcutOnSystemStartMenu) {
        $shortcutFiles += "$SystemStartMenu\$($shortcut.Name).lnk"
    }
}

# Initialize a flag to indicate if all shortcuts exist
$allShortcutsExist = $true

# Check if each shortcut exists
foreach ($shortcutFile in $shortcutFiles) {
    if (!(Test-Path -Path $shortcutFile)) {
        Write-Host "Shortcut $(Split-Path -Leaf $shortcutFile) wasn't detected at $(Split-Path -Parent $shortcutFile)."
        $allShortcutsExist = $false
    }
}

# Exit with a status code based on whether all shortcuts exist
if ($allShortcutsExist) {
    Write-Host "All shortcuts were detected."
    Exit 0
} else {
    Write-Host "Not all shortcuts were detected."
    Exit 1
}