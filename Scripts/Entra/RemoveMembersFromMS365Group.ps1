# This scripr remove all members from MS 365 group leaving owners 

Connect-MgGraph -Scopes GroupMember.ReadWrite.All, Group.Read.All

$GroupId = "GROUP ID HERE7"

# Get group owners
Write-Host "Fetching group owners..."
$owners = Get-MgGroupOwner -GroupId $GroupId | Select-Object -ExpandProperty Id

# Get group members
Write-Host "Fetching group members..."
$members = Get-MgGroupMember -GroupId $GroupId -All | Select-Object -ExpandProperty Id

# Remove non-owner members with retry logic and batching
$batchSize = 100  # Increased batch size
$retryDelay = 1   # Reduced delay in seconds before retrying
$maxRetries = 3   # Maximum retries per member
$totalMembers = $members.Count
$removedCount = 0

foreach ($memberId in $members) {
    if ($owners -notcontains $memberId) {
        $retryCount = 0
        $removed = $false

        while ($retryCount -lt $maxRetries -and -not $removed) {
            try {
                Write-Host "Removing member: $memberId (Attempt: $($retryCount + 1))"
                Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $memberId
                $removed = $true
                $removedCount++
                Write-Host "Successfully removed: $memberId ($removedCount/$totalMembers)"
            } catch {
                Write-Host "Failed to remove $memberId. Error: $_.Exception.Message"
                $retryCount++
                Start-Sleep -Seconds $retryDelay
            }
        }

        if (-not $removed) {
            Write-Host "Skipping member $memberId after $maxRetries failed attempts."
        }

        # Pause every batch to prevent throttling
        if ($removedCount % $batchSize -eq 0) {
            Write-Host "Pausing for 5 seconds to prevent throttling..."  # Reduced pause time
            Start-Sleep -Seconds 5
        }
    }
}

Write-Host "All non-owner members processed."