# This script adds users to SharePoint site groups (Owners or Members) for multiple sites, based on information in a CSV file. It handles errors and logs any failed operations.
# CSV File: Contains user assignments; each row should have SiteURL, UserEmail, and Role (e.g., Owner or Member).
# It requries to install and configure PnP module and give it a correct API permissions for AppRegistration


# Import the PnP PowerShell Module

# Define the input and output paths
$csvFilePath = "C:\XXX.csv"
$errorLogFilePath = "C:\AAA.txt"

# AAD (Entra ID) App Registration Details
$tenantId = "XXX"       # Replace with your Tenant ID
$clientId = "XXX"       # Replace with your Client ID

# Clear previous error log (if any)
if (Test-Path $errorLogFilePath) {
    Remove-Item $errorLogFilePath -Force
}

# Import users and groups from the CSV
$data = Import-Csv -Path $csvFilePath

foreach ($entry in $data) {
    $siteUrl = $entry.SiteURL
    $userEmail = $entry.UserEmail
    $role = $entry.Role

    try {
        # Interactive login with Azure App Registration
        Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Interactive

        # Determine the group based on the role
        if ($role -eq "Owner") {
            $groupName = "Owners"  # Default SharePoint Owners Group
        } elseif ($role -eq "Member") {
            $groupName = "Members"  # Default SharePoint Members Group
        } else {
            # Unsupported role; log and skip
            $message = "Skipped user ${userEmail} at ${siteUrl} - unsupported role '${role}'"
            Write-Warning $message
            Add-Content -Path $errorLogFilePath -Value $message
            continue
        }

        # Get the SharePoint Group
        Write-Host "Fetching group: $groupName from site $siteUrl"
        $group = Get-PnPGroup | Where-Object { $_.Title -like "*$groupName" }
        if (-not $group) {
            throw "Cannot find group '${groupName}' in site: ${siteUrl}"
        }

        # Add the user to the SharePoint Group using Add-PnPGroupMember
        Write-Host "Adding $userEmail to group: $($group.Title)"
        Add-PnPGroupMember -LoginName $userEmail -Group $group.Title
        Write-Host "Successfully added ${userEmail} to '${groupName}' group in ${siteUrl}"

    } catch {
        # Log the error
        $errorMessage = "Error for user ${userEmail} in site ${siteUrl}: $($_.Exception.Message)"
        Write-Warning $errorMessage
        Add-Content -Path $errorLogFilePath -Value $errorMessage
    }
}