# Ask user if they want to generate a CSV template
$generateTemplate = Read-Host -Prompt "Do you want to generate a CSV template? (y/n)"

if ($generateTemplate -eq "y" -or $generateTemplate -eq "Y" -or $generateTemplate -eq "yes" -or $generateTemplate -eq "Yes") {
    # Get the current script directory
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $templatePath = Join-Path $scriptDir "teams_channel_users_template.csv"
    
    # Create sample CSV template
    $templateContent = @"
userPrincipalName,isUser
john.doe@company.com,TRUE
jane.smith@company.com,TRUE
guest@external.com,FALSE
"@
    
    try {
        $templateContent | Out-File -FilePath $templatePath -Encoding UTF8
        Write-Host "CSV template created successfully at: $templatePath" -ForegroundColor Green
        Write-Host "Template includes sample data. Please modify it with your actual user data." -ForegroundColor Yellow
        Write-Host "CSV format: userPrincipalName,isUser" -ForegroundColor Cyan
        Write-Host "  - userPrincipalName: Email address of the user" -ForegroundColor Cyan
        Write-Host "  - isUser: TRUE for regular users, FALSE for guests/external users" -ForegroundColor Cyan
    } catch {
        Write-Error "Failed to create CSV template. Error details: $_"
    }
    
    # Ask if user wants to use the template they just created
    $useTemplate = Read-Host -Prompt "Do you want to use this template file? (y/n)"
    if ($useTemplate -eq "y" -or $useTemplate -eq "Y" -or $useTemplate -eq "yes" -or $useTemplate -eq "Yes") {
        $csvPath = $templatePath
        Write-Host "Using template file: $csvPath" -ForegroundColor Green
    } else {
        $csvPath = Read-Host -Prompt "Enter the path to your CSV file"
    }
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