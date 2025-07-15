# Prompt for the CSV path, team ID, and channel name
$csvPath = Read-Host -Prompt "Enter the path to your CSV file"
$teamId = Read-Host -Prompt "Enter your team ID"
$channelName = Read-Host -Prompt "Enter your channel name"

# Check if the specified CSV file exists
if (!(Test-Path -Path $csvPath)) {
    Write-Error "The specified CSV file does not exist: $csvPath. Please check the file path and try again."
    break
}

# Import the CSV file
try {
    $users = Import-Csv -Path $csvPath
    Write-Host "CSV file successfully imported." -ForegroundColor Green
} catch {
    Write-Error "Failed to import the CSV file: $csvPath. Error details: $_"
    break
}

# Connect to Microsoft Teams
try {
    Connect-MicrosoftTeams
    Write-Host "Successfully connected to Microsoft Teams." -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to Microsoft Teams. Please ensure that the Microsoft Teams module is installed and that you're logged in. Error details: $_"
    break
}

# Loop through each user and add them to the channel
foreach ($user in $users) {
    try {
        if ($user.isUser -eq "TRUE") {
            # Add the user to the channel
            Add-TeamChannelUser -GroupId $teamId -DisplayName $channelName -User $user.userPrincipalName
            Write-Host "Successfully added user $($user.userPrincipalName) to the channel '$channelName'." -ForegroundColor Green
        } else {
            Write-Host "Skipping non-user: $($user.userPrincipalName)" -ForegroundColor Yellow
        }
    } catch {
        Write-Error "Failed to add user $($user.userPrincipalName) to the channel '$channelName'. Error details: $_"
    }
}

Write-Host "Script finished." -ForegroundColor Green