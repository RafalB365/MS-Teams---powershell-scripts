# Creates a MS Teams shared channels based on CSV file containing at least two columns: ChannelName and Email.

# Define your team ID
$teamId = "TYPE-TEAM-ID"  # Replace with the actual Team ID
$outputCsv = "C:\temp\channelsIDs_export.csv"
$logFile = "C:\temp\teams_creation_log.txt"

# Read CSV file
$channelData = Import-Csv -Path $outputCsv
$existingChannels = @{}

# Initialize log file
"Script started on $(Get-Date)`n" | Out-File -Append $logFile

# Step 1: Get existing channels in the team
$existingTeamsChannels = Get-TeamChannel -GroupId $teamId
foreach ($channel in $existingTeamsChannels) {
    $existingChannels[$channel.DisplayName] = $channel.Id
}

# Step 2: Process CSV data
foreach ($entry in $channelData) {
    $channelName = $entry.ChannelName
    $userEmail = $entry.Email

    # Check if channel already exists, if not, create it
    if ($existingChannels.ContainsKey($channelName)) {
        $channelId = $existingChannels[$channelName]
        Write-Output "Channel '$channelName' already exists."
    } else {
        Write-Output "Creating shared channel: $channelName"
        try {
            $channel = New-TeamChannel -GroupId $teamId -DisplayName $channelName -MembershipType Shared
            $channelId = $channel.Id
            $existingChannels[$channelName] = $channelId
            "Created shared channel: $channelName (ID: $channelId)" | Out-File -Append $logFile
        } catch {
            "Error creating channel '$channelName': $_" | Out-File -Append $logFile
            continue
        }
    }

    # Step 3: Find user in Microsoft 365
    try {
        $user = Get-MgUser -Filter "mail eq '$userEmail'"
        if ($user) {
            Write-Output "Adding user $userEmail to channel $channelName"
            Add-TeamChannelUser -GroupId $teamId -DisplayName $channelName -User $user.UserPrincipalName
            "User $userEmail added to channel $channelName" | Out-File -Append $logFile
        } else {
            Write-Output "User not found: $userEmail"
            "Skipping: User not found ($userEmail)" | Out-File -Append $logFile
        }
    } catch {
        "Error adding user '$userEmail' to '$channelName': $_" | Out-File -Append $logFile
        continue
    }
}

# Final log entry
"Script completed on $(Get-Date)`n" | Out-File -Append $logFile
Write-Output "Process completed. Check log file at: $logFile"
