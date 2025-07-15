# This script will allow members to create Teams in MS Teams if the MS 365 group creation is blocked on the org wide settings
# Import required modules
Import-Module Microsoft.Graph.Beta.Identity.DirectoryManagement
Import-Module Microsoft.Graph.Beta.Groups

# Note: MgGraph connection is handled in the main script

# Get the group name from user input
$GroupName = Read-Host "Enter the name for the Security Group that will be allowed to create Teams"

# Create the security group
Write-Host "Creating security group: $GroupName" -ForegroundColor Green
try {
    $newGroup = New-MgBetaGroup -DisplayName $GroupName -SecurityEnabled -MailEnabled:$false -MailNickname ($GroupName -replace '\s', '')
    $GroupObjectId = $newGroup.Id
    Write-Host "Security group '$GroupName' created successfully with ID: $GroupObjectId" -ForegroundColor Green
} catch {
    Write-Error "Failed to create security group: $($_.Exception.Message)"
    exit 1
}

# Define variables
$AllowGroupCreation = "False"

# Check if the "Group.Unified" setting exists
$settingsObjectID = (Get-MgBetaDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ).Id

# If the setting doesn't exist, create it
if (!$settingsObjectID) {
    $params = @{
        templateId = "62375ab9-6b52-47ed-826b-58e47e0e304b"
        values = @(
            @{
                name  = "EnableMSStandardBlockedWords"
                value = "true"
            }
        )
    }
    New-MgBetaDirectorySetting -BodyParameter $params

    # Re-fetch the settings ID
    $settingsObjectID = (Get-MgBetaDirectorySetting | Where-Object -Property DisplayName -Value "Group.Unified" -EQ).Id
}

# Update the settings with the provided Object ID
$params = @{
    templateId = "62375ab9-6b52-47ed-826b-58e47e0e304b"
    values = @(
        @{
            name  = "EnableGroupCreation"
            value = $AllowGroupCreation
        }
        @{
            name  = "GroupCreationAllowedGroupId"
            value = $GroupObjectId  # Use the provided Object ID
        }
    )
}

# Apply the updated settings
Update-MgBetaDirectorySetting -DirectorySettingId $settingsObjectID -BodyParameter $params

# Output the updated settings for verification
(Get-MgBetaDirectorySetting -DirectorySettingId $settingsObjectID).Values

# Final message to user
Write-Host "`n=== CONFIGURATION COMPLETE ===" -ForegroundColor Yellow
Write-Host "Only members of the '$GroupName' security group will be allowed to create Teams in MS Teams." -ForegroundColor Cyan
Write-Host "To allow users to create Teams, add them to the '$GroupName' security group." -ForegroundColor Cyan
Write-Host "Group Object ID: $GroupObjectId" -ForegroundColor White
