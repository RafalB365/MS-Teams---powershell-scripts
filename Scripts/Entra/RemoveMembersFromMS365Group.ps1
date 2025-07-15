# This script removes all members from MS 365 group leaving owners 

param(
    [Parameter(Mandatory = $true)]
    [string]$GroupId,
    
    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 100,
    
    [Parameter(Mandatory = $false)]
    [int]$RetryDelay = 1,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 3,
    
    [Parameter(Mandatory = $false)]
    [string]$LogFile = $null
)

# Function to write log messages
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    Write-Host $logMessage
    
    if ($LogFile) {
        Add-Content -Path $LogFile -Value $logMessage
    }
}

# Function to validate Group ID format
function Test-GroupId {
    param([string]$Id)
    
    # Check if it's a valid GUID format
    try {
        $guid = [System.Guid]::Parse($Id)
        return $true
    } catch {
        return $false
    }
}

# Validate Group ID
if (-not (Test-GroupId -Id $GroupId)) {
    Write-Log "Invalid Group ID format. Please provide a valid GUID." "ERROR"
    exit 1
}

Write-Log "Starting MS365 Group member removal process for Group ID: $GroupId"
Write-Log "Configuration - Batch Size: $BatchSize, Retry Delay: $RetryDelay seconds, Max Retries: $MaxRetries"

# Connect to Microsoft Graph
try {
    Write-Log "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes GroupMember.ReadWrite.All, Group.Read.All -ErrorAction Stop
    Write-Log "Successfully connected to Microsoft Graph"
} catch {
    Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" "ERROR"
    exit 1
}

# Verify group exists
try {
    Write-Log "Verifying group exists..."
    $group = Get-MgGroup -GroupId $GroupId -ErrorAction Stop
    Write-Log "Group found: $($group.DisplayName)"
} catch {
    Write-Log "Group not found or access denied: $($_.Exception.Message)" "ERROR"
    exit 1
}

# Get group owners
Write-Log "Fetching group owners..."
try {
    $owners = Get-MgGroupOwner -GroupId $GroupId | Select-Object -ExpandProperty Id
    Write-Log "Found $($owners.Count) owner(s)"
} catch {
    Write-Log "Failed to fetch group owners: $($_.Exception.Message)" "ERROR"
    exit 1
}

# Get group members
Write-Log "Fetching group members..."
try {
    $members = Get-MgGroupMember -GroupId $GroupId -All | Select-Object -ExpandProperty Id
    Write-Log "Found $($members.Count) member(s)"
} catch {
    Write-Log "Failed to fetch group members: $($_.Exception.Message)" "ERROR"
    exit 1
}

# Calculate non-owner members
$nonOwnerMembers = $members | Where-Object { $owners -notcontains $_ }
Write-Log "Found $($nonOwnerMembers.Count) non-owner member(s) to remove"

if ($nonOwnerMembers.Count -eq 0) {
    Write-Log "No non-owner members to remove. Exiting."
    exit 0
}

# Remove non-owner members with retry logic and batching
$totalMembers = $nonOwnerMembers.Count
$removedCount = 0
$failedRemovals = @()
$skippedCount = 0

foreach ($memberId in $nonOwnerMembers) {
    $retryCount = 0
    $removed = $false

    while ($retryCount -lt $MaxRetries -and -not $removed) {
        try {
            Write-Log "Removing member: $memberId (Attempt: $($retryCount + 1))"
            Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $memberId -ErrorAction Stop
            $removed = $true
            $removedCount++
            Write-Log "Successfully removed: $memberId ($removedCount/$totalMembers)"
        } catch {
            $errorMessage = $_.Exception.Message
            Write-Log "Failed to remove $memberId. Error: $errorMessage" "WARNING"
            
            # Check if it's a permanent error (member doesn't exist, etc.)
            if ($errorMessage -like "*NotFound*" -or $errorMessage -like "*does not exist*") {
                Write-Log "Member $memberId appears to no longer exist. Skipping." "WARNING"
                $skippedCount++
                break
            }
            
            $retryCount++
            if ($retryCount -lt $MaxRetries) {
                Write-Log "Retrying in $RetryDelay seconds..."
                Start-Sleep -Seconds $RetryDelay
            }
        }
    }

    if (-not $removed) {
        Write-Log "Failed to remove member $memberId after $MaxRetries attempts." "ERROR"
        $failedRemovals += $memberId
    }

    # Pause every batch to prevent throttling
    if ($removedCount % $BatchSize -eq 0 -and $removedCount -gt 0) {
        Write-Log "Pausing for 5 seconds to prevent throttling..."
        Start-Sleep -Seconds 5
    }
}

# Summary
Write-Log "=== REMOVAL SUMMARY ===" 
Write-Log "Total non-owner members found: $totalMembers"
Write-Log "Successfully removed: $removedCount"
Write-Log "Skipped (not found): $skippedCount"
Write-Log "Failed to remove: $($failedRemovals.Count)"

if ($failedRemovals.Count -gt 0) {
    Write-Log "Failed member IDs:" "ERROR"
    foreach ($failedId in $failedRemovals) {
        Write-Log "  - $failedId" "ERROR"
    }
}

Write-Log "All non-owner members processed."

# Disconnect from Microsoft Graph
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Log "Disconnected from Microsoft Graph"
} catch {
    Write-Log "Warning: Could not disconnect from Microsoft Graph: $($_.Exception.Message)" "WARNING"
}