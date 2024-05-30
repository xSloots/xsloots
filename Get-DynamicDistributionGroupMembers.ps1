<#
    .SYNOPSIS
        This script is used to get members of a Dynamic Distribution Group in Exchange Online and display them in a grid view or export them to a CSV file.
    .DESCRIPTION
        You need to have proper rights to execute the script with minumum of Exchange Recipient Administrator role.
        The script will prompt you to select a Dynamic Distribution Group from a list of all Dynamic Distribution Groups in Exchange Online.
        You can then choose to display the members of the selected group in a grid view or export them to a CSV file.
    .NOTES
        File Name      : Get-DynamicDistributionGroupMembers.ps1
        Author         : Lars Sloots
        Prerequisite   : Exchange Online Management Module
        Version        : 1.0
        Date           : 30/05/2024
#>

# Load the .NET assembly for Windows Forms. This is used to create the input box and the checkboxes.
Add-Type -AssemblyName System.Windows.Forms

# Check if the ExchangeOnlineManagement module is installed. If not, install it.
if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
    }

# Connect to Exchange Online without showing the banner.
Connect-ExchangeOnline -ShowBanner:$false

# Get all Dynamic Distribution Groups and display them in an Out-GridView for the user to select one.
$group = Get-DynamicDistributionGroup | Select-Object Name | Out-GridView -PassThru

# Create the input box where the user can select the options for displaying the results.
$inputBox = New-Object System.Windows.Forms.Form 
$inputBox.Text = 'Input Required'
$inputBox.Size = New-Object System.Drawing.Size(300,150) 
$inputBox.StartPosition = 'CenterScreen'

# Create the OK button for the input box.
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(10,70)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$inputBox.AcceptButton = $okButton
$inputBox.Controls.Add($okButton)

# Create the checkbox for the grid view option. If this is checked, the results will be displayed in a grid view.
$gridViewBox = New-Object System.Windows.Forms.CheckBox 
$gridViewBox.Location = New-Object System.Drawing.Point(10,10)
$gridViewBox.Size = New-Object System.Drawing.Size(260,20) 
$gridViewBox.Text = 'Use gridview to display the results'
$inputBox.Controls.Add($gridViewBox) 

# Create the checkbox for the CSV export option. If this is checked, the results will be exported to a CSV file.
$csvExportBox = New-Object System.Windows.Forms.CheckBox 
$csvExportBox.Location = New-Object System.Drawing.Point(10,40)
$csvExportBox.Size = New-Object System.Drawing.Size(260,20) 
$csvExportBox.Text = 'Export to CSV'
$inputBox.Controls.Add($csvExportBox) 

# Show the input box and wait for the user to click the OK button.
$inputBox.Topmost = $true
$inputBox.Add_Shown({$inputBox.Activate()})
$result = $inputBox.ShowDialog()

# If the user clicked the OK button, get the selected Dynamic Distribution Group and its members, and display them according to the selected options.
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $name = $group.Name
    $useGridView = $gridViewBox.Checked
    $csvExport = $csvExportBox.Checked

    # Get the Dynamic Distribution Group and its members.
    $recipients = Get-Recipient -RecipientPreviewFilter (Get-DynamicDistributionGroup -Identity $name).RecipientFilter | Select-Object DisplayName, Title, Department
    if ($useGridView) {
        $recipients | Out-GridView
    } else {
        $recipients | Format-Table
    }
    if ($csvExport) {
        $name = $group.Name -replace '[\\/:*?"<>|]', '_' # Remove invalid characters from the name
        $csvPath = "C:\Temp\$name.csv"
            # Check if the C:\Temp directory exists
            if (!(Test-Path -Path "C:\Temp")) {
            # If the directory doesn't exist, create it
            New-Item -ItemType Directory -Path "C:\Temp"
            }
        $recipients | Export-Csv -Path $csvPath -Delimiter ';' -NoTypeInformation
        Invoke-Item $csvPath
    }
}

# Disconnect from Exchange Online without asking for confirmation.
Disconnect-ExchangeOnline -Confirm:$false 
