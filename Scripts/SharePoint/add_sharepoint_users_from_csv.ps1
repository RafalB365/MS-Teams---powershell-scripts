# This script adds users to SharePoint site groups (Owners or Members) for multiple sites, based on information in a CSV file. It handles errors and logs any failed operations.
# CSV File: Contains user assignments; each row should have SiteURL, UserEmail, and Role (e.g., Owner or Member).
# PnP connection configuration should be set up in the main menu PnP configurations.
# Compatible with both PnP.PowerShell (PowerShell 7+) and SharePointPnPPowerShellOnline (PowerShell 5.1)

# Function to detect and import the appropriate PnP module
function Import-PnPModule {
    $psVersion = $PSVersionTable.PSVersion
    
    # Check for modern PnP.PowerShell first
    $modernPnP = Get-Module -Name PnP.PowerShell -ListAvailable -ErrorAction SilentlyContinue
    $legacyPnP = Get-Module -Name SharePointPnPPowerShellOnline -ListAvailable -ErrorAction SilentlyContinue
    
    if ($modernPnP -and $psVersion.Major -ge 7) {
        Write-Host "Using PnP.PowerShell (modern module)" -ForegroundColor Green
        Import-Module PnP.PowerShell -ErrorAction Stop
        return "Modern"
    }
    elseif ($legacyPnP) {
        Write-Host "Using SharePointPnPPowerShellOnline (legacy module)" -ForegroundColor Yellow
        Import-Module SharePointPnPPowerShellOnline -ErrorAction Stop
        return "Legacy"
    }
    elseif ($modernPnP -and $psVersion.Major -lt 7) {
        Write-Host "Warning: PnP.PowerShell detected but may not work properly on PowerShell $($psVersion.Major).$($psVersion.Minor)" -ForegroundColor Yellow
        Write-Host "Attempting to use PnP.PowerShell anyway..." -ForegroundColor Yellow
        Import-Module PnP.PowerShell -ErrorAction Stop
        return "Modern"
    }
    else {
        throw "No PnP module found. Please install PnP.PowerShell (PowerShell 7+) or SharePointPnPPowerShellOnline (PowerShell 5.1)"
    }
}

# Function to connect to SharePoint based on module type
function Connect-SharePoint {
    param(
        [string]$SiteUrl,
        [string]$ModuleType
    )
    
    if ($ModuleType -eq "Modern") {
        # Use modern PnP.PowerShell connection
        Connect-PnPOnline -Url $SiteUrl -Interactive
    }
    else {
        # Use legacy SharePointPnPPowerShellOnline connection
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin
    }
}

# Try to import the appropriate PnP module
try {
    $moduleType = Import-PnPModule
    Write-Host "Successfully loaded PnP module" -ForegroundColor Green
}
catch {
    Write-Host "Error loading PnP module: $_" -ForegroundColor Red
    Write-Host "Please install the appropriate PnP module:" -ForegroundColor Yellow
    Write-Host "  - For PowerShell 7+: Install-Module PnP.PowerShell -Scope CurrentUser" -ForegroundColor Yellow
    Write-Host "  - For PowerShell 5.1: Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

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
    Write-Host "Please edit the file with your actual site URLs, user emails, and roles" -ForegroundColor Yellow
    Write-Host "Supported roles: Owner, Member, Visitor" -ForegroundColor Cyan
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
        # Connect using the appropriate PnP module
        Write-Host "Connecting to SharePoint site: $siteUrl" -ForegroundColor Cyan
        Connect-SharePoint -SiteUrl $siteUrl -ModuleType $moduleType

        # Determine the group based on the role
        if ($role -eq "Owner") {
            $groupName = "Owners"  # Default SharePoint Owners Group
        } elseif ($role -eq "Member") {
            $groupName = "Members"  # Default SharePoint Members Group
        } elseif ($role -eq "Visitor") {
            $groupName = "Visitors"  # Default SharePoint Visitors Group
        } else {
            # Unsupported role; log and skip
            $message = "Skipped user ${userEmail} at ${siteUrl} - unsupported role '${role}' (supported: Owner, Member, Visitor)"
            Write-Warning $message
            Add-Content -Path $errorLogFilePath -Value $message
            continue
        }

        # Get the SharePoint Group
        Write-Host "Fetching group: $groupName from site $siteUrl" -ForegroundColor Gray
        $group = Get-PnPGroup | Where-Object { $_.Title -like "*$groupName" }
        if (-not $group) {
            throw "Cannot find group '${groupName}' in site: ${siteUrl}"
        }

        # Add the user to the SharePoint Group using Add-PnPGroupMember
        Write-Host "Adding $userEmail to group: $($group.Title)" -ForegroundColor Gray
        Add-PnPGroupMember -LoginName $userEmail -Group $group.Title
        Write-Host "✓ Successfully added ${userEmail} to '${groupName}' group in ${siteUrl}" -ForegroundColor Green

    } catch {
        # Log the error
        $errorMessage = "Error for user ${userEmail} in site ${siteUrl}: $($_.Exception.Message)"
        Write-Warning $errorMessage
        Add-Content -Path $errorLogFilePath -Value $errorMessage
    }
}

Write-Host ""
Write-Host "Script execution completed!" -ForegroundColor Green
Write-Host "Module used: $(if ($moduleType -eq 'Modern') { 'PnP.PowerShell' } else { 'SharePointPnPPowerShellOnline' })" -ForegroundColor Cyan
if (Test-Path $errorLogFilePath) {
    Write-Host "Check error log at: $errorLogFilePath" -ForegroundColor Yellow
} else {
    Write-Host "All operations completed successfully!" -ForegroundColor Green
}