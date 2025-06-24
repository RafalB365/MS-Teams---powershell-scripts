#This script removes an user from all channels in a single, specified Team

# Import the Microsoft Teams PowerShell module
Import-Module MicrosoftTeams

# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Define the team and user
$teamId = "TEAM-ID"
$userId = "USER UPN"

# Get all channels in the team
$channels = Get-TeamChannel -GroupId $teamId

# Loop through each channel and remove the user
foreach ($channel in $channels) {
    try {
        Remove-TeamChannelUser -GroupId $teamId -DisplayName $channel.DisplayName -User $userId
        Write-Output "Removed user $userId from channel $($channel.DisplayName)"
    } catch {
        Write-Output "Failed to remove user $userId from channel $($channel.DisplayName): $_"
    }
}

# Disconnect from Microsoft Teams
Disconnect-MicrosoftTeams