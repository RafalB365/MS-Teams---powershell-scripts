# This script will allow members to create Teams in MS Teams if the MS 365 group creation is blocked on the org wide settings
# Import required modules
Import-Module Microsoft.Graph.Beta.Identity.DirectoryManagement
Import-Module Microsoft.Graph.Beta.Groups

# Connect to Microsoft Graph with the necessary permissions
Connect-MgGraph -Scopes "Directory.ReadWrite.All", "Group.Read.All"

# Define variables
$GroupObjectId = "XXX"  # Replace with your Group Object ID
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
