# This script bulk-creates groups in Azure AD from a CSV, reporting on all success/failure outcomes, and carefully handles both Microsoft 365 and Security groups.
# Input Required: Assumes the CSV contains at least displayName, mail, and groupType columns.


# Import the CSV file
$groups = Import-Csv -Path "C:\XXX.csv"

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
        New-MgGroup @groupParams
        Write-Host "Group '$($group.displayName)' created successfully."
        $successGroups += $group.displayName
    }
    catch {
        Write-Host "Failed to create group '$($group.displayName)': $($_.Exception.Message)" -ForegroundColor Red
        $failedGroups += @{
            GroupName = $group.displayName
            Error     = $_.Exception.Message
            Params    = $groupParams  # Log the parameters that caused the error
        }
    }
}

# Output summary after the process
Write-Host "`nGroup Creation Summary:"
Write-Host "-----------------------"

# Log successful groups
if ($successGroups.Count -gt 0) {
    Write-Host "`nSuccessfully Created Groups:"
    foreach ($group in $successGroups) {
        Write-Host "- $group" -ForegroundColor Green
    }
} else {
    Write-Host "No groups were created successfully." -ForegroundColor Yellow
}

# Log failed groups with errors
if ($failedGroups.Count -gt 0) {
    Write-Host "`nFailed to Create the Following Groups:"
    foreach ($failure in $failedGroups) {
        Write-Host "- Group: $($failure.GroupName)" -ForegroundColor Red
        Write-Host "  Error: $($failure.Error)" -ForegroundColor DarkYellow
        Write-Host "  Params: $($failure.Params | ConvertTo-Json -Depth 5)" -ForegroundColor DarkGray  # Log the parameters
    }
} else {
    Write-Host "No groups failed to create." -ForegroundColor Green
}
