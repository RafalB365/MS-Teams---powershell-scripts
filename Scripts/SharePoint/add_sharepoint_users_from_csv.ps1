# This script adds users to SharePoint site groups (Owners or Members) for multiple sites, based on information in a CSV file. It handles errors and logs any failed operations.
# CSV File: Contains user assignments; each row should have SiteURL, UserEmail, and Role (e.g., Owner or Member).
# PnP connection configuration should be set up in the main menu PnP configurations.

# Import the PnP PowerShell Module
Import-Module PnP.PowerShell -ErrorAction SilentlyContinue

# Define the working directory and file paths
# Use the application work folder from the main menu configuration
$appFolderPath = Join-Path -Path $env:USERPROFILE -ChildPath "MS 365 Admin Alembic"
$csvFilePath = Join-Path -Path $appFolderPath -ChildPath "SharePoint_Users_Import.csv"
$errorLogFilePath = Join-Path -Path $appFolderPath -ChildPath "SharePoint_Users_Import_Errors.txt"

# Ensure the work folder exists
if (-not (Test-Path -Path $appFolderPath -PathType Container)) {
    New-Item -Path $appFolderPath -ItemType Directory -Force | Out-Null
}

# Function to generate template CSV
function New-TemplateCSV {
    param([string]$FilePath)
    
    $templateData = @(
        [PSCustomObject]@{
            SiteURL = "https://yourtenant.sharepoint.com/sites/site1"
            UserEmail = "user1@yourdomain.com"
            Role = "Owner"
        },
        [PSCustomObject]@{
            SiteURL = "https://yourtenant.sharepoint.com/sites/site1"
            UserEmail = "user2@yourdomain.com"
            Role = "Member"
        },
        [PSCustomObject]@{
            SiteURL = "https://yourtenant.sharepoint.com/sites/site2"
            UserEmail = "user3@yourdomain.com"
            Role = "Visitor"
        }
    )
    
    $templateData | Export-Csv -Path $FilePath -NoTypeInformation
    Write-Host "Template CSV file created at: $FilePath" -ForegroundColor Green
    Write-Host "Please edit the file with your actual site URLs, user emails, and roles (Owner/Member)" -ForegroundColor Yellow
}

# Ask user if they want to generate a template CSV
Write-Host "SharePoint Users Import Script" -ForegroundColor Cyan
Write-Host "=============================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $csvFilePath)) {
    $generateTemplate = Read-Host "CSV file not found. Would you like to generate a template CSV file? (Y/N)"
    
    if ($generateTemplate -eq 'Y' -or $generateTemplate -eq 'y') {
        New-TemplateCSV -FilePath $csvFilePath
        Write-Host ""
        Write-Host "Please edit the template file and run the script again." -ForegroundColor Yellow
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        # Exit the script after generating the template CSV file
        return
        Write-Host "Please provide the path to your CSV file:" -ForegroundColor Yellow
        $userCsvPath = Read-Host "CSV file path"
        
        if (Test-Path $userCsvPath) {
            $csvFilePath = $userCsvPath
        } else {
            Write-Host "File not found. Please check the path and try again." -ForegroundColor Red
            return
        }
    }
}

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
        # Connect using PnP settings from main menu configuration
        # Note: Ensure PnP connection is configured in the main menu before running this script
        Connect-PnPOnline -Url $siteUrl -Interactive

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

Write-Host ""
Write-Host "Script execution completed!" -ForegroundColor Green
if (Test-Path $errorLogFilePath) {
    Write-Host "Check error log at: $errorLogFilePath" -ForegroundColor Yellow
} else {
    Write-Host "All operations completed successfully!" -ForegroundColor Green
}