#This script removes a user from all channels in a single, specified Team
#Requires -Module MicrosoftTeams

<#
.SYNOPSIS
Removes a user from all channels in a specified Microsoft Teams team.

.DESCRIPTION
This script removes a specified user from all channels within a Microsoft Teams team.
It assumes that the Microsoft Teams PowerShell module is already loaded and connected.
The script will prompt for the Team ID and User UPN interactively.

.EXAMPLE
.\removeuserfromateamchannels.ps1
#>

# Prompt for Team ID and User UPN
$TeamId = Read-Host "Enter the Team ID (Group ID)"
$UserId = Read-Host "Enter the User UPN (e.g., user@domain.com)"

# Validate input
if ([string]::IsNullOrWhiteSpace($TeamId)) {
    Write-Error "Team ID cannot be empty"
    exit 1
}

if ([string]::IsNullOrWhiteSpace($UserId)) {
    Write-Error "User UPN cannot be empty"
    exit 1
}

Write-Output "Team ID: $TeamId"
Write-Output "User UPN: $UserId"
Write-Output ""

# Validate that we're connected to Microsoft Teams
try {
    $null = Get-Team -GroupId $TeamId -ErrorAction Stop
    Write-Output "Successfully validated connection to Microsoft Teams and team access"
} catch {
    Write-Error "Failed to access team or not connected to Microsoft Teams: $_"
    exit 1
}

# Get all channels in the team
try {
    $channels = Get-TeamChannel -GroupId $TeamId -ErrorAction Stop
    Write-Output "Found $($channels.Count) channels in the team"
} catch {
    Write-Error "Failed to retrieve channels for team ${TeamId}: $_"
    exit 1
}

# Loop through each channel and remove the user
$successCount = 0
$errorCount = 0

foreach ($channel in $channels) {
    try {
        Remove-TeamChannelUser -GroupId $TeamId -DisplayName $channel.DisplayName -User $UserId -ErrorAction Stop
        Write-Output "✓ Removed user $UserId from channel '$($channel.DisplayName)'"
        $successCount++
    } catch {
        Write-Warning "✗ Failed to remove user $UserId from channel '$($channel.DisplayName)': $_"
        $errorCount++
    }
}

# Summary
Write-Output "`nSummary:"
Write-Output "Successfully removed user from $successCount channels"
if ($errorCount -gt 0) {
    Write-Output "Failed to remove user from $errorCount channels"
}