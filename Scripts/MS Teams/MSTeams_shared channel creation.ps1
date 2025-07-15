# Creates a MS Teams shared channels based on CSV file containing at least two columns: ChannelName and Email.

# Define your team ID
$teamId = "TYPE-TEAM-ID"  # Replace with the actual Team ID

# Get the Alembic work folder path
$appFolderPath = Join-Path -Path $env:USERPROFILE -ChildPath "MS 365 Admin Alembic"

# Prompt for CSV file path
$csvPath = Read-Host "Enter the path to your CSV file (or press Enter for default: $appFolderPath\teams_channels_template.csv)"
if ([string]::IsNullOrWhiteSpace($csvPath)) {
    $csvPath = Join-Path -Path $appFolderPath -ChildPath "teams_channels_template.csv"
}

# Read CSV file
$channelData = Import-Csv -Path $csvPath
$existingChannels = @{}

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
        } catch {
            Write-Error "Error creating channel '$channelName': $_"
            continue
        }
    }

    # Step 3: Find user in Microsoft 365
    try {
        $user = Get-MgUser -Filter "mail eq '$userEmail'"
        if ($user) {
            Write-Output "Adding user $userEmail to channel $channelName"
            Add-TeamChannelUser -GroupId $teamId -DisplayName $channelName -User $user.UserPrincipalName
        } else {
            Write-Output "User not found: $userEmail"
        }
    } catch {
        Write-Error "Error adding user '$userEmail' to '$channelName': $_"
        continue
    }
}

Write-Output "Process completed."
