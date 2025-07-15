# This script bulk-creates groups in Azure AD from a CSV, reporting on all success/failure outcomes, and carefully handles both Microsoft 365 and Security groups.
# Input Required: Assumes the CSV contains at least displayName, mail, and groupType columns.

# Get the MS 365 Admin Alembic work folder
$appName = "MS 365 Admin Alembic"
$appFolderPath = Join-Path -Path $env:USERPROFILE -ChildPath $appName
$csvFilePath = Join-Path -Path $appFolderPath -ChildPath "groups_to_create.csv"

# Create directory if it doesn't exist
if (-not (Test-Path -Path $appFolderPath -PathType Container)) {
    New-Item -Path $appFolderPath -ItemType Directory -Force | Out-Null
}

Write-Host "=== Azure AD Group Bulk Creation ===" -ForegroundColor Green
Write-Host ""

# Module management and authentication
Write-Host "Checking and loading required modules..." -ForegroundColor Yellow

# Remove any existing Microsoft Graph modules to avoid conflicts
$graphModules = Get-Module -Name "Microsoft.Graph*" -ErrorAction SilentlyContinue
if ($graphModules) {
    Write-Host "Removing existing Microsoft Graph modules to avoid conflicts..." -ForegroundColor Yellow
    Remove-Module -Name "Microsoft.Graph*" -Force -ErrorAction SilentlyContinue
}

# Import required modules
try {
    Import-Module Microsoft.Graph.Groups -Force -ErrorAction Stop
    Write-Host "✓ Microsoft.Graph.Groups module loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "✗ Failed to load Microsoft.Graph.Groups module: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please install the module using: Install-Module Microsoft.Graph.Groups -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

# Check if user is connected to Microsoft Graph
try {
    $context = Get-MgContext -ErrorAction Stop
    if ($null -eq $context) {
        Write-Host "Not connected to Microsoft Graph. Please connect first." -ForegroundColor Red
        Write-Host "Use: Connect-MgGraph -Scopes 'Group.ReadWrite.All','Directory.ReadWrite.All'" -ForegroundColor Yellow
        exit 1
    }
    Write-Host "✓ Connected to Microsoft Graph as: $($context.Account)" -ForegroundColor Green
}
catch {
    Write-Host "Not connected to Microsoft Graph. Attempting to connect..." -ForegroundColor Yellow
    try {
        Connect-MgGraph -Scopes 'Group.ReadWrite.All','Directory.ReadWrite.All' -ErrorAction Stop
        Write-Host "✓ Successfully connected to Microsoft Graph" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please run: Connect-MgGraph -Scopes 'Group.ReadWrite.All','Directory.ReadWrite.All'" -ForegroundColor Yellow
        exit 1
    }
}

Write-Host ""

# Ask user if they want to generate a CSV template
$generateCsv = Read-Host "Do you want to generate a CSV template file? (y/n)"

if ($generateCsv -eq 'y' -or $generateCsv -eq 'Y') {
    # Generate CSV template with sample data
    $csvTemplate = @"
displayName,mail,groupType
Marketing Team,marketing@company.com,Microsoft 365
Sales Security Group,,Security
Finance Team,finance@company.com,Microsoft 365
IT Security Group,,Security
HR Team,hr@company.com,Microsoft 365
"@
    
    # Save CSV template to work folder
    $csvTemplate | Out-File -FilePath $csvFilePath -Encoding UTF8
    
    Write-Host "CSV template generated at: $csvFilePath" -ForegroundColor Green
    Write-Host ""
    Write-Host "Please edit the CSV file with your group information and then run this script again." -ForegroundColor Yellow
    Write-Host "CSV contains the following columns:" -ForegroundColor Gray
    Write-Host "- displayName: The display name of the group" -ForegroundColor Gray
    Write-Host "- mail: Email address for Microsoft 365 groups (leave empty for Security groups)" -ForegroundColor Gray
    Write-Host "- groupType: Either 'Microsoft 365' or 'Security'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Opening the CSV file for editing..." -ForegroundColor Yellow
    
    # Open the CSV file for editing
    Start-Process notepad.exe -ArgumentList $csvFilePath
    
    Write-Host "Press any key when you've finished editing the CSV file..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
else {
    Write-Host "Please place your CSV file in the work folder: $appFolderPath" -ForegroundColor Yellow
    Write-Host "The CSV file should be named 'groups_to_create.csv'" -ForegroundColor Yellow
    Write-Host "Required columns: displayName, mail, groupType" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Press any key when you've placed the CSV file..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

# Check if CSV file exists
if (-not (Test-Path -Path $csvFilePath)) {
    Write-Host "ERROR: CSV file not found at $csvFilePath" -ForegroundColor Red
    Write-Host "Please ensure the CSV file exists and try again." -ForegroundColor Red
    exit 1
}

Write-Host "Found CSV file. Processing groups..." -ForegroundColor Green
Write-Host ""

# Import the CSV file
$groups = Import-Csv -Path $csvFilePath

# Validate CSV data
if (-not $groups -or $groups.Count -eq 0) {
    Write-Host "ERROR: No data found in CSV file or CSV is empty" -ForegroundColor Red
    exit 1
}

# Check if required columns exist
$requiredColumns = @('displayName', 'groupType')
$csvColumns = $groups[0].PSObject.Properties.Name
$missingColumns = $requiredColumns | Where-Object { $_ -notin $csvColumns }

if ($missingColumns.Count -gt 0) {
    Write-Host "ERROR: Missing required columns in CSV: $($missingColumns -join ', ')" -ForegroundColor Red
    Write-Host "Required columns: $($requiredColumns -join ', ')" -ForegroundColor Yellow
    exit 1
}

Write-Host "CSV validation passed. Found $($groups.Count) groups to process." -ForegroundColor Green
Write-Host ""

# Initialize arrays to track success and failures
$successGroups = @()
$failedGroups = @()

foreach ($group in $groups) {
    # Check if mail is present, otherwise fall back to a sanitized displayName
    if (-not [string]::IsNullOrEmpty($group.mail)) {
        $mailNickname = $group.mail.Split('@')[0]
    }
    else {
        # Fallback: use the displayName for mailNickname
        $mailNickname = $group.displayName -replace '[^a-zA-Z0-9]', '-'
    }

    # Further sanitize the mailNickname (remove invalid characters, make it lowercase, and limit to 64 chars)
    $mailNickname = $mailNickname.ToLower() -replace '[^a-zA-Z0-9]', '-' 
    $mailNickname = $mailNickname.Substring(0, [Math]::Min($mailNickname.Length, 64))  # Limit to 64 characters

    # Define group parameters with null checks
    $groupParams = @{
        DisplayName  = $group.displayName
        MailNickname = $mailNickname
    }

    # Check groupType to create Microsoft 365 group or Security group
    if ($group.groupType -eq "Microsoft 365") {
        # Create Microsoft 365 group (Unified)
        $groupParams.MailEnabled = $true
        $groupParams.SecurityEnabled = $false
        $groupParams.GroupTypes = "Unified"
        $groupParams.Visibility = "Private"  # You can adjust this as needed
    }
    elseif ($group.groupType -eq "Security") {
        # Create Security group
        $groupParams.SecurityEnabled = $true
        $groupParams.MailEnabled = $false  # Security groups can't be mail-enabled
    }
    else {
        Write-Host "Skipping group '$($group.displayName)': Invalid groupType." -ForegroundColor Yellow
        continue  # Skip to the next iteration if groupType is not valid
    }

    # Log group parameters before attempting to create the group
    Write-Host "Attempting to create group with parameters: $($groupParams | ConvertTo-Json -Depth 5)"

    # Try to add group to Microsoft 365
    try {
        # Check if group with same display name already exists
        $existingGroup = Get-MgGroup -Filter "displayName eq '$($group.displayName)'" -ErrorAction SilentlyContinue
        if ($existingGroup) {
            Write-Host "Group '$($group.displayName)' already exists. Skipping..." -ForegroundColor Yellow
            continue
        }

        # Create the group
        $newGroup = New-MgGroup @groupParams -ErrorAction Stop
        Write-Host "Group '$($group.displayName)' created successfully. ID: $($newGroup.Id)" -ForegroundColor Green
        $successGroups += $group.displayName
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Failed to create group '$($group.displayName)': $errorMessage" -ForegroundColor Red
        
        # Check for specific errors and provide helpful suggestions
        if ($errorMessage -like "*already exists*") {
            Write-Host "  Suggestion: A group with this name or mail nickname already exists" -ForegroundColor Yellow
        }
        elseif ($errorMessage -like "*insufficient privileges*") {
            Write-Host "  Suggestion: Check that you have Group.ReadWrite.All and Directory.ReadWrite.All permissions" -ForegroundColor Yellow
        }
        elseif ($errorMessage -like "*mailNickname*") {
            Write-Host "  Suggestion: The mail nickname '$mailNickname' may be invalid or already in use" -ForegroundColor Yellow
        }
        
        $failedGroups += @{
            GroupName = $group.displayName
            Error     = $errorMessage
            Params    = $groupParams  # Log the parameters that caused the error
        }
    }
}

# Output summary after the process
Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host "           GROUP CREATION SUMMARY" -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "Total groups processed: $($groups.Count)" -ForegroundColor White
Write-Host "Successfully created: $($successGroups.Count)" -ForegroundColor Green
Write-Host "Failed to create: $($failedGroups.Count)" -ForegroundColor Red
Write-Host "===================================================" -ForegroundColor Cyan

# Log successful groups
if ($successGroups.Count -gt 0) {
    Write-Host "`n✓ Successfully Created Groups:" -ForegroundColor Green
    foreach ($group in $successGroups) {
        Write-Host "  • $group" -ForegroundColor Green
    }
} else {
    Write-Host "`n⚠ No groups were created successfully." -ForegroundColor Yellow
}

# Log failed groups with errors
if ($failedGroups.Count -gt 0) {
    Write-Host "`n✗ Failed to Create the Following Groups:" -ForegroundColor Red
    foreach ($failure in $failedGroups) {
        Write-Host "  • Group: $($failure.GroupName)" -ForegroundColor Red
        Write-Host "    Error: $($failure.Error)" -ForegroundColor DarkYellow
        Write-Host "    Params: $($failure.Params | ConvertTo-Json -Depth 5)" -ForegroundColor DarkGray
        Write-Host ""
    }
} else {
    Write-Host "`n✓ All groups were created successfully!" -ForegroundColor Green
}

Write-Host "===================================================" -ForegroundColor Cyan
