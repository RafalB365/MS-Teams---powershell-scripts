# MS 365 Admin Alembic with Encrypted Storage and Module Management
# Version 1.0
$toolkitVersion = "1.0"

# Update application name and folder structure
$appName = "MS 365 Admin Alembic"
$appFolderName = $appName
$appFolderPath = Join-Path -Path $env:USERPROFILE -ChildPath $appFolderName
$logPath = Join-Path -Path $appFolderPath -ChildPath "toolkit.log"

# Handle migration from old folder structure if necessary
$oldFolderPath = Join-Path -Path $env:USERPROFILE -ChildPath ".m365toolkit"
if (Test-Path -Path $oldFolderPath -PathType Container) {
    if (-not (Test-Path -Path $appFolderPath -PathType Container)) {
        # Create new directory and migrate content
        New-Item -Path $appFolderPath -ItemType Directory -Force | Out-Null
        Copy-Item -Path "$oldFolderPath\*" -Destination $appFolderPath -Recurse -Force
        Write-Host "Migrated settings and cache from old location to $appFolderPath" -ForegroundColor Yellow
    }
}

# Create directory if it doesn't exist
if (-not (Test-Path -Path $appFolderPath -PathType Container)) {
    New-Item -Path $appFolderPath -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logPath -Append -Encoding utf8
}

Clear-Host

# ASCII Art Header
Write-Host ""
Write-Host "  ███╗   ███╗███████╗██████╗  ██████╗ ███████╗" -ForegroundColor DarkGreen
Write-Host "  ████╗ ████║██╔════╝╚════██╗██╔════╝ ██╔════╝" -ForegroundColor DarkGreen
Write-Host "  ██╔████╔██║███████╗ █████╔╝███████╗ ███████╗" -ForegroundColor DarkGreen
Write-Host "  ██║╚██╔╝██║╚════██║ ╚═══██╗██╔═══██╗╚════██║" -ForegroundColor DarkGreen
Write-Host "  ██║ ╚═╝ ██║███████║██████╔╝╚██████╔╝███████║" -ForegroundColor DarkGreen
Write-Host "  ╚═╝     ╚═╝╚══════╝╚═════╝  ╚═════╝ ╚══════╝" -ForegroundColor DarkGreen
Write-Host ""
Write-Host "    █████╗ ██████╗ ███╗   ███╗██╗███╗   ██╗" -ForegroundColor DarkGreen
Write-Host "   ██╔══██╗██╔══██╗████╗ ████║██║████╗  ██║" -ForegroundColor DarkGreen
Write-Host "   ███████║██║  ██║██╔████╔██║██║██╔██╗ ██║" -ForegroundColor DarkGreen
Write-Host "   ██╔══██║██║  ██║██║╚██╔╝██║██║██║╚██╗██║" -ForegroundColor DarkGreen
Write-Host "   ██║  ██║██████╔╝██║ ╚═╝ ██║██║██║ ╚████║" -ForegroundColor DarkGreen
Write-Host "   ╚═╝  ╚═╝╚═════╝ ╚═╝     ╚═╝╚═╝╚═╝  ╚═══╝" -ForegroundColor DarkGreen
Write-Host ""
Write-Host "    █████╗ ██╗     ███████╗███╗   ███╗██████╗ ██╗ ██████╗" -ForegroundColor DarkGreen
Write-Host "   ██╔══██╗██║     ██╔════╝████╗ ████║██╔══██╗██║██╔════╝" -ForegroundColor DarkGreen
Write-Host "   ███████║██║     █████╗  ██╔████╔██║██████╔╝██║██║     " -ForegroundColor DarkGreen
Write-Host "   ██╔══██║██║     ██╔══╝  ██║╚██╔╝██║██╔══██╗██║██║     " -ForegroundColor DarkGreen
Write-Host "   ██║  ██║███████╗███████╗██║ ╚═╝ ██║██████╔╝██║╚██████╗" -ForegroundColor DarkGreen
Write-Host "   ╚═╝  ╚═╝╚══════╝╚══════╝╚═╝     ╚═╝╚═════╝ ╚═╝ ╚═════╝" -ForegroundColor DarkGreen
Write-Host ""
Write-Host "   ╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor DarkGreen
Write-Host "   ║                    v$toolkitVersion - PowerShell Toolkit      ║" -ForegroundColor DarkGreen
Write-Host "   ║              Your MS365 Administration Swiss Army Knife!      ║" -ForegroundColor DarkGreen
Write-Host "   ╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor DarkGreen
Write-Host ""

# Create desktop shortcut to the application folder
function Create-DesktopShortcut {
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path -Path $desktopPath -ChildPath "$appName.lnk"
    
    if (-not (Test-Path -Path $shortcutPath)) {
        try {
            $WshShell = New-Object -ComObject WScript.Shell
            $shortcut = $WshShell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = $appFolderPath
            $shortcut.Description = "MS 365 Admin Alembic Configuration Folder"
            $shortcut.IconLocation = "shell32.dll,4"  # Folder icon
            $shortcut.Save()
            Write-Host "Desktop shortcut created successfully." -ForegroundColor Green
        }
        catch {
            Write-Log "Error creating desktop shortcut: $_"
            Write-Host "Error creating desktop shortcut: $_" -ForegroundColor Red
        }
    }
}

# Script information with categories and descriptions
$scriptInfo = @{
    # SharePoint scripts
    "1" = @{
        Name = "Add SharePoint Users from CSV"
        Description = "Adds users to SharePoint site groups (Owners or Members) for multiple sites based on CSV file"
        Url = "local"
        ScriptName = "add_sharepoint_users_from_csv.ps1"
        LocalPath = "Scripts\SharePoint\add_sharepoint_users_from_csv.ps1"
        RequiredModules = @("PnP.PowerShell")
        RequiredScopes = @("Sites.FullControl.All")
        Category = "SharePoint"
    }
    
    # MS Teams scripts
    "2" = @{
        Name = "Allow Teams Creation"
        Description = "Allow members to create Teams in MS Teams if the MS 365 group creation is blocked on the org wide settings"
        Url = "local"
        ScriptName = "allowteamscreation.ps1"
        LocalPath = "Scripts\MS Teams\allowteamscreation.ps1"
        RequiredModules = @("Microsoft.Graph.Beta.Identity.DirectoryManagement", "Microsoft.Graph.Beta.Groups")
        RequiredScopes = @("Group.ReadWrite.All", "Directory.ReadWrite.All")
        Category = "MS Teams"
    }
    "3" = @{
        Name = "Import Users from CSV to Private Teams Channel"
        Description = "Imports users from CSV file to a private Teams channel with template generation"
        Url = "local"
        ScriptName = "importusersfromcsvtoprivateteamschannel.ps1"
        LocalPath = "Scripts\MS Teams\importusersfromcsvtoprivateteamschannel.ps1"
        RequiredModules = @("MicrosoftTeams")
        RequiredScopes = @("TeamMember.ReadWrite.All", "Group.ReadWrite.All")
        Category = "MS Teams"
    }
    "4" = @{
        Name = "Save all Teams owners to CSV"
        Description = "Since there is no way to see the owners of other Teams, this script lists them all and saves to CSV"
        Url = "local"
        ScriptName = "listallteamowners.ps1"
        LocalPath = "Scripts\MS Teams\listallteamowners.ps1"
        RequiredModules = @("Microsoft.Graph.Users")
        RequiredScopes = @("GroupMember.Read.All", "Group.Read.All", "User.Read.All")
        Category = "MS Teams"
    }
    "5" = @{
        Name = "MS Teams Shared Channel Creation"
        Description = "Creates MS Teams shared channels based on CSV file containing ChannelName and Email columns"
        Url = "local"
        ScriptName = "MSTeams_shared channel creation.ps1"
        LocalPath = "Scripts\MS Teams\MSTeams_shared channel creation.ps1"
        RequiredModules = @("MicrosoftTeams", "Microsoft.Graph.Users")
        RequiredScopes = @("TeamMember.ReadWrite.All", "Group.ReadWrite.All", "User.Read.All")
        Category = "MS Teams"
    }
    "6" = @{
        Name = "Remove User from Team Channels"
        Description = "Removes a user from all channels in a single, specified Team"
        Url = "local"
        ScriptName = "removeuserfromateamchannels.ps1"
        LocalPath = "Scripts\MS Teams\removeuserfromateamchannels.ps1"
        RequiredModules = @("MicrosoftTeams")
        RequiredScopes = @("TeamMember.ReadWrite.All", "Group.ReadWrite.All")
        Category = "MS Teams"
    }
    
    # Exchange scripts
    "7" = @{
        Name = "Create DL based on MS 365 group"
        Description = "This script creates an Exchange Distribution List based on the MS 365 group members"
        Url = "local"
        ScriptName = "createDL.ps1"
        LocalPath = "Scripts\Exchange\createDL.ps1"
        RequiredModules = @("ExchangeOnlineManagement")
        RequiredScopes = @()  # Exchange Online uses its own authentication
        Category = "Exchange"
    }
    
    # Entra (Azure AD) scripts
    "8" = @{
        Name = "Bulk Group Creation from CSV"
        Description = "Bulk creates Microsoft 365 and Security groups from a CSV file with interactive CSV generation"
        Url = "local"
        ScriptName = "bulkgroupcreationfromCSV.ps1"
        LocalPath = "Scripts\Entra\bulkgroupcreationfromCSV.ps1"
        RequiredModules = @("Microsoft.Graph.Groups")
        RequiredScopes = @("Group.ReadWrite.All", "Directory.ReadWrite.All")
        Category = "Entra"
    }
    "9" = @{
        Name = "Remove Members from MS365 Group"
        Description = "Removes all members from MS 365 group leaving owners with batch processing and retry logic"
        Url = "local"
        ScriptName = "RemoveMembersFromMS365Group.ps1"
        LocalPath = "Scripts\Entra\RemoveMembersFromMS365Group.ps1"
        RequiredModules = @("Microsoft.Graph.Groups")
        RequiredScopes = @("GroupMember.ReadWrite.All", "Group.Read.All")
        Category = "Entra"
    }
}

# Global settings with secure PnP connection parameters
$global:Settings = @{
    UseCache = $true
    DefaultParameters = @{}
    PnPConnection = @{
        RememberCredentials = $true
    }
}

# Initialize paths
$settingsPath = $env:M365_TOOLKIT_SETTINGS_PATH
if (-not $settingsPath) {
    $settingsPath = Join-Path -Path $appFolderPath -ChildPath "settings.xml"
}
$cacheDir = $env:M365_TOOLKIT_CACHE_DIR
if (-not $cacheDir) {
    $cacheDir = $appFolderPath
}
$cacheDir = $appFolderPath

# Function to encrypt a string
function Protect-String {
    param ([Parameter(Mandatory=$true)][string]$PlainText)
    
    if ([string]::IsNullOrEmpty($PlainText)) { return $null }
    
    try {
        $secureString = ConvertTo-SecureString -String $PlainText -AsPlainText -Force
        $encrypted = ConvertFrom-SecureString -SecureString $secureString
        return $encrypted
    } catch {
        Write-Host "Error encrypting string: $_" -ForegroundColor Red
        return $null
    }
}

# Function to decrypt an encrypted string
function Unprotect-String {
    param ([Parameter(Mandatory=$true)][string]$EncryptedText)
    
    if ([string]::IsNullOrEmpty($EncryptedText)) { return $null }
    
    try {
        $secureString = ConvertTo-SecureString -String $EncryptedText
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
        $plainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        return $plainText
    } catch {
        Write-Host "Error decrypting string: $_" -ForegroundColor Red
        return $null
    }
}

# Function to load saved settings
function Load-Settings {
    if (Test-Path $settingsPath) {
        try {
            $savedSettings = Import-Clixml -Path $settingsPath
            
            # Copy basic settings
            $global:Settings.UseCache = $savedSettings.UseCache
            $global:Settings.DefaultParameters = $savedSettings.DefaultParameters
            $global:Settings.PnPConnection.RememberCredentials = $savedSettings.PnPConnection.RememberCredentials
            
            # Handle encrypted credentials
            if ($savedSettings.PnPConnection.ContainsKey("EncryptedSiteUrl")) {
                $global:Settings.PnPConnection.EncryptedSiteUrl = $savedSettings.PnPConnection.EncryptedSiteUrl
                if ($savedSettings.PnPConnection.EncryptedSiteUrl) {
                    $global:Settings.PnPConnection.SiteUrl = Unprotect-String $savedSettings.PnPConnection.EncryptedSiteUrl
                }
            }
            
            if ($savedSettings.PnPConnection.ContainsKey("EncryptedClientId")) {
                $global:Settings.PnPConnection.EncryptedClientId = $savedSettings.PnPConnection.EncryptedClientId
                if ($savedSettings.PnPConnection.EncryptedClientId) {
                    $global:Settings.PnPConnection.ClientId = Unprotect-String $savedSettings.PnPConnection.EncryptedClientId
                }
            }
            
            Write-Host "Settings loaded from $settingsPath" -ForegroundColor Gray
        }
        catch {
            Write-Log "Error: $_"
            Write-Host "Could not load settings file: $_" -ForegroundColor Yellow
        }
    }
    else {
        # Initialize directory if it doesn't exist
        $settingsDir = Split-Path -Path $settingsPath -Parent
        if (-not (Test-Path -Path $settingsDir)) {
            New-Item -Path $settingsDir -ItemType Directory | Out-Null
        }
        Save-Settings
    }
}

# Function to save settings with encryption
function Save-Settings {
    try {
        # Create a copy of settings to modify for saving
        $settingsToSave = @{
            UseCache = $global:Settings.UseCache
            DefaultParameters = $global:Settings.DefaultParameters
            PnPConnection = @{
                RememberCredentials = $global:Settings.PnPConnection.RememberCredentials
            }
        }
        
        # Encrypt sensitive data before saving
        if ($global:Settings.PnPConnection.ContainsKey("SiteUrl") -and 
            -not [string]::IsNullOrEmpty($global:Settings.PnPConnection.SiteUrl)) {
            $settingsToSave.PnPConnection.EncryptedSiteUrl = Protect-String $global:Settings.PnPConnection.SiteUrl
        }
        
        if ($global:Settings.PnPConnection.ContainsKey("ClientId") -and 
            -not [string]::IsNullOrEmpty($global:Settings.PnPConnection.ClientId)) {
            $settingsToSave.PnPConnection.EncryptedClientId = Protect-String $global:Settings.PnPConnection.ClientId
        }
        
        # Save to file
        $settingsToSave | Export-Clixml -Path $settingsPath -Force
        Write-Host "Settings saved securely to $settingsPath" -ForegroundColor Gray
    }
    catch {
        Write-Log "Error: $_"
        Write-Host "Failed to save settings: $_" -ForegroundColor Red
    }
}

# Function to show the categorized menu
function Show-Menu {
    Write-Host "`nSelect an option:" -ForegroundColor Yellow
    
    # Define custom category order
    $categoryOrder = @("SharePoint", "MS Teams", "Exchange", "Entra")
    
    foreach ($category in $categoryOrder) {
        # Check if category exists in scripts
        $categoryExists = $scriptInfo.Values | Where-Object { $_.Category -eq $category }
        if ($categoryExists) {
            Write-Host "`n[$category]" -ForegroundColor Cyan
            $categoryScripts = $scriptInfo.GetEnumerator() | Where-Object { $_.Value.Category -eq $category } | Sort-Object Name
            
            foreach ($script in $categoryScripts) {
                Write-Host "$($script.Key). $($script.Value.Name)"
            }
        }
    }
    
    Write-Host "`n[Configuration]" -ForegroundColor Cyan
    Write-Host "C. Configure Settings"
    Write-Host "D. Check Dependencies"
    Write-Host "P. Configure PnP Connection"
    Write-Host "F. Open Configuration Folder"
    Write-Host "0. Exit`n"
}

# Function to show script description
function Show-ScriptDescription {
    param ([string]$ScriptKey)
    
    if ($scriptInfo.ContainsKey($ScriptKey)) {
        $script = $scriptInfo[$ScriptKey]
        Write-Host "`n=== $($script.Name) ===" -ForegroundColor Green
        Write-Host "Description: $($script.Description)" -ForegroundColor Gray
        Write-Host "Required Modules: $($script.RequiredModules -join ', ')" -ForegroundColor Gray
        Write-Host "Required Scopes: $($script.RequiredScopes -join ', ')" -ForegroundColor Gray
        Write-Host ""
    }
}

# Function to execute script from GitHub or local file
function Execute-ScriptFromGitHub {
    param (
        [string]$ScriptUrl,
        [string]$ScriptName,
        [array]$RequiredModules,
        [array]$RequiredScopes,
        [bool]$UseCache = $true,
        [string]$LocalPath = ""
    )
    
    try {
        if ($ScriptUrl -eq "local" -and -not [string]::IsNullOrEmpty($LocalPath)) {
            # Execute local script
            Write-Host "Executing local script $ScriptName..." -ForegroundColor Yellow
            
            # Get the directory where the main script is located
            $scriptDir = Split-Path -Path $MyInvocation.PSCommandPath -Parent
            $fullLocalPath = Join-Path -Path $scriptDir -ChildPath $LocalPath
            
            if (Test-Path -Path $fullLocalPath) {
                $scriptContent = Get-Content -Path $fullLocalPath -Raw
                Write-Host "Found local script at: $fullLocalPath" -ForegroundColor Gray
            } else {
                throw "Local script not found at: $fullLocalPath"
            }
        } else {
            # Download script from GitHub
            Write-Host "Downloading and executing $ScriptName..." -ForegroundColor Yellow
            $scriptContent = Invoke-RestMethod -Uri $ScriptUrl -ErrorAction Stop
        }
        
        # Save to cache if enabled
        if ($UseCache) {
            $cacheFile = Join-Path -Path $cacheDir -ChildPath $ScriptName
            $scriptContent | Out-File -FilePath $cacheFile -Encoding UTF8
            Write-Host "Script cached to $cacheFile" -ForegroundColor Gray
        }
        
        # Execute the script
        Invoke-Expression $scriptContent
        
        Write-Host "Script execution completed." -ForegroundColor Green
    }
    catch {
        Write-Log "Error executing script ${ScriptName}: $_"
        Write-Host "Error executing script: $_" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

# Function to configure settings
function Configure-Settings {
    Write-Host "`n=== Configure Settings ===" -ForegroundColor Green
    
    # Cache settings
    Write-Host "`nCurrent cache setting: $($global:Settings.UseCache)" -ForegroundColor Gray
    $cacheChoice = Read-Host "Enable caching? (y/n)"
    $global:Settings.UseCache = ($cacheChoice -eq 'y' -or $cacheChoice -eq 'Y')
    
    # PnP Connection settings
    Write-Host "`nCurrent remember credentials setting: $($global:Settings.PnPConnection.RememberCredentials)" -ForegroundColor Gray
    $credChoice = Read-Host "Remember PnP credentials? (y/n)"
    $global:Settings.PnPConnection.RememberCredentials = ($credChoice -eq 'y' -or $credChoice -eq 'Y')
    
    # Save settings
    Save-Settings
    Write-Host "Settings saved successfully." -ForegroundColor Green
    
    Write-Host "`nPress any key to return to menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

# Function to check all dependencies
function Check-AllDependencies {
    Write-Host "`n=== Checking Dependencies ===" -ForegroundColor Green
    
    # Get all unique modules from scripts
    $allModules = $scriptInfo.Values | ForEach-Object { $_.RequiredModules } | Select-Object -Unique | Sort-Object
    
    Write-Host "`n--- Checking PowerShell Modules ---" -ForegroundColor Cyan
    
    foreach ($module in $allModules) {
        Write-Host "Checking $module..." -NoNewline
        
        try {
            $installedModule = Get-Module -Name $module -ListAvailable -ErrorAction SilentlyContinue
            if ($installedModule) {
                Write-Host " ✓ Installed (v$($installedModule[0].Version))" -ForegroundColor Green
            } else {
                Write-Host " ✗ Not installed" -ForegroundColor Red
                $installChoice = Read-Host "Install $module? (y/n)"
                if ($installChoice -eq 'y' -or $installChoice -eq 'Y') {
                    Write-Host "Installing $module..." -ForegroundColor Yellow
                    Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
                    Write-Host "$module installed successfully." -ForegroundColor Green
                }
            }
        }
        catch {
            Write-Host " ✗ Error checking module: $_" -ForegroundColor Red
        }
    }
    
    Write-Host "`n--- Checking Service Connections ---" -ForegroundColor Cyan
    
    # Check Microsoft Graph connection
    Write-Host "Checking Microsoft Graph connection..." -NoNewline
    try {
        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($mgContext) {
            Write-Host " ✓ Connected as $($mgContext.Account)" -ForegroundColor Green
            Write-Host "  Current scopes: $($mgContext.Scopes -join ', ')" -ForegroundColor Gray
            
            # Check if all required scopes are present
            $allRequiredScopes = $scriptInfo.Values | ForEach-Object { $_.RequiredScopes } | Select-Object -Unique | Sort-Object
            $missingScopes = $allRequiredScopes | Where-Object { $_ -notin $mgContext.Scopes }
            
            if ($missingScopes.Count -gt 0) {
                Write-Host "  ⚠ Missing scopes: $($missingScopes -join ', ')" -ForegroundColor Yellow
                $reconnectChoice = Read-Host "Reconnect with all required scopes? (y/n)"
                if ($reconnectChoice -eq 'y' -or $reconnectChoice -eq 'Y') {
                    Write-Host "Connecting to Microsoft Graph with all required scopes..." -ForegroundColor Yellow
                    try {
                        Connect-MgGraph -Scopes $allRequiredScopes -ErrorAction Stop
                        Write-Host "✓ Connected to Microsoft Graph with all required scopes." -ForegroundColor Green
                    }
                    catch {
                        Write-Host "✗ Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
                        Write-Host "  Try connecting manually with: Connect-MgGraph -Scopes '$($allRequiredScopes -join "','")'" -ForegroundColor Yellow
                    }
                }
            }
        } else {
            Write-Host " ✗ Not connected" -ForegroundColor Red
            $connectChoice = Read-Host "Connect to Microsoft Graph? (y/n)"
            if ($connectChoice -eq 'y' -or $connectChoice -eq 'Y') {
                $allRequiredScopes = $scriptInfo.Values | ForEach-Object { $_.RequiredScopes } | Select-Object -Unique | Sort-Object
                Write-Host "Connecting to Microsoft Graph with scopes: $($allRequiredScopes -join ', ')" -ForegroundColor Yellow
                try {
                    Connect-MgGraph -Scopes $allRequiredScopes -ErrorAction Stop
                    Write-Host "✓ Connected to Microsoft Graph successfully." -ForegroundColor Green
                }
                catch {
                    Write-Host "✗ Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
                    Write-Host "  Try connecting manually with: Connect-MgGraph -Scopes '$($allRequiredScopes -join "','")'" -ForegroundColor Yellow
                }
            }
        }
    }
    catch {
        Write-Host " ✗ Error checking Microsoft Graph connection: $_" -ForegroundColor Red
    }
    
    # Check Microsoft Teams connection
    Write-Host "Checking Microsoft Teams connection..." -NoNewline
    try {
        $teamsContext = Get-CsTenant -ErrorAction SilentlyContinue
        if ($teamsContext) {
            Write-Host " ✓ Connected to tenant: $($teamsContext.DisplayName)" -ForegroundColor Green
        } else {
            Write-Host " ✗ Not connected" -ForegroundColor Red
            $connectChoice = Read-Host "Connect to Microsoft Teams? (y/n)"
            if ($connectChoice -eq 'y' -or $connectChoice -eq 'Y') {
                Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Yellow
                try {
                    Connect-MicrosoftTeams -ErrorAction Stop
                    Write-Host "✓ Connected to Microsoft Teams successfully." -ForegroundColor Green
                }
                catch {
                    Write-Host "✗ Failed to connect to Microsoft Teams: $_" -ForegroundColor Red
                }
            }
        }
    }
    catch {
        Write-Host " ✗ Error checking Microsoft Teams connection: $_" -ForegroundColor Red
        $connectChoice = Read-Host "Connect to Microsoft Teams? (y/n)"
        if ($connectChoice -eq 'y' -or $connectChoice -eq 'Y') {
            try {
                Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Yellow
                Connect-MicrosoftTeams -ErrorAction Stop
                Write-Host "✓ Connected to Microsoft Teams successfully." -ForegroundColor Green
            }
            catch {
                Write-Host "✗ Failed to connect to Microsoft Teams: $_" -ForegroundColor Red
            }
        }
    }
    
    # Check Exchange Online connection
    Write-Host "Checking Exchange Online connection..." -NoNewline
    try {
        # First check if the ExchangeOnlineManagement module is imported
        $exchangeModule = Get-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
        if (-not $exchangeModule) {
            try {
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
            }
            catch {
                Write-Host " ✗ ExchangeOnlineManagement module not available" -ForegroundColor Red
                Write-Host "  Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Yellow
                return
            }
        }
        
        # Check if we're connected by trying to get connection info
        $connectionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($connectionInfo) {
            Write-Host " ✓ Connected to Exchange Online ($($connectionInfo.Organization))" -ForegroundColor Green
        } else {
            Write-Host " ✗ Not connected" -ForegroundColor Red
            $connectChoice = Read-Host "Connect to Exchange Online? (y/n)"
            if ($connectChoice -eq 'y' -or $connectChoice -eq 'Y') {
                Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
                try {
                    Connect-ExchangeOnline -ErrorAction Stop
                    Write-Host "✓ Connected to Exchange Online successfully." -ForegroundColor Green
                }
                catch {
                    Write-Host "✗ Failed to connect to Exchange Online: $_" -ForegroundColor Red
                    Write-Host "  Try connecting manually with: Connect-ExchangeOnline" -ForegroundColor Yellow
                }
            }
        }
    }
    catch {
        Write-Host " ✗ Error checking Exchange Online connection: $_" -ForegroundColor Red
        $connectChoice = Read-Host "Connect to Exchange Online? (y/n)"
        if ($connectChoice -eq 'y' -or $connectChoice -eq 'Y') {
            try {
                Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
                Connect-ExchangeOnline -ErrorAction Stop
                Write-Host "✓ Connected to Exchange Online successfully." -ForegroundColor Green
            }
            catch {
                Write-Host "✗ Failed to connect to Exchange Online: $_" -ForegroundColor Red
                Write-Host "  Try connecting manually with: Connect-ExchangeOnline" -ForegroundColor Yellow
            }
        }
    }
    
    Write-Host "`n--- Connection Summary ---" -ForegroundColor Cyan
    
    # Final status check
    try {
        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($mgContext) {
            Write-Host "✓ Microsoft Graph: Connected as $($mgContext.Account)" -ForegroundColor Green
        } else {
            Write-Host "✗ Microsoft Graph: Not connected" -ForegroundColor Red
        }
    } catch {
        Write-Host "✗ Microsoft Graph: Connection error" -ForegroundColor Red
    }
    
    try {
        $teamsContext = Get-CsTenant -ErrorAction SilentlyContinue
        if ($teamsContext) {
            Write-Host "✓ Microsoft Teams: Connected to $($teamsContext.DisplayName)" -ForegroundColor Green
        } else {
            Write-Host "✗ Microsoft Teams: Not connected" -ForegroundColor Red
        }
    } catch {
        Write-Host "✗ Microsoft Teams: Connection error" -ForegroundColor Red
    }
    
    try {
        $connectionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($connectionInfo) {
            Write-Host "✓ Exchange Online: Connected to $($connectionInfo.Organization)" -ForegroundColor Green
        } else {
            Write-Host "✗ Exchange Online: Not connected" -ForegroundColor Red
        }
    } catch {
        Write-Host "✗ Exchange Online: Connection error" -ForegroundColor Red
    }
    
    Write-Host "`nDependency and connection check completed." -ForegroundColor Green
    Write-Host "`nPress any key to return to menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

# Function to configure PnP connection
function Configure-PnPConnection {
    Write-Host "`n=== Configure PnP Connection ===" -ForegroundColor Green
    
    # Site URL
    $currentSiteUrl = if ($global:Settings.PnPConnection.ContainsKey("SiteUrl")) { $global:Settings.PnPConnection.SiteUrl } else { "Not set" }
    Write-Host "Current Site URL: $currentSiteUrl" -ForegroundColor Gray
    $siteUrl = Read-Host "Enter SharePoint Site URL (or press Enter to keep current)"
    if (-not [string]::IsNullOrEmpty($siteUrl)) {
        $global:Settings.PnPConnection.SiteUrl = $siteUrl
    }
    
    # Client ID
    $currentClientId = if ($global:Settings.PnPConnection.ContainsKey("ClientId")) { $global:Settings.PnPConnection.ClientId } else { "Not set" }
    Write-Host "Current Client ID: $currentClientId" -ForegroundColor Gray
    $clientId = Read-Host "Enter App Registration Client ID (or press Enter to keep current)"
    if (-not [string]::IsNullOrEmpty($clientId)) {
        $global:Settings.PnPConnection.ClientId = $clientId
    }
    
    # Save settings
    Save-Settings
    Write-Host "PnP connection settings saved successfully." -ForegroundColor Green
    
    Write-Host "`nPress any key to return to menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

# Function to open configuration folder
function Open-ConfigFolder {
    Write-Host "`n=== Opening Configuration Folder ===" -ForegroundColor Green
    try {
        Start-Process explorer.exe -ArgumentList $appFolderPath
        Write-Host "Configuration folder opened: $appFolderPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Error opening configuration folder: $_" -ForegroundColor Red
    }
    
    Write-Host "`nPress any key to return to menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

# === MAIN EXECUTION ===
Load-Settings
Create-DesktopShortcut

$exitScript = $false
do {
    Show-Menu
    $choice = Read-Host "Enter your choice"

    switch ($choice.ToUpper()) {
        # Script options
        "1" {
            Show-ScriptDescription "1"
            $script = $scriptInfo["1"]
            if ($script.ContainsKey("LocalPath")) {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache -LocalPath $script.LocalPath
            } else {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
            }
        }
        "2" {
            Show-ScriptDescription "2"
            $script = $scriptInfo["2"]
            if ($script.ContainsKey("LocalPath")) {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache -LocalPath $script.LocalPath
            } else {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
            }
        }
        "3" {
            Show-ScriptDescription "3"
            $script = $scriptInfo["3"]
            if ($script.ContainsKey("LocalPath")) {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache -LocalPath $script.LocalPath
            } else {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
            }
        }
        "4" {
            Show-ScriptDescription "4"
            $script = $scriptInfo["4"]
            if ($script.ContainsKey("LocalPath")) {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache -LocalPath $script.LocalPath
            } else {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
            }
        }
        "5" {
            Show-ScriptDescription "5"
            $script = $scriptInfo["5"]
            if ($script.ContainsKey("LocalPath")) {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache -LocalPath $script.LocalPath
            } else {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
            }
        }
        "6" {
            Show-ScriptDescription "6"
            $script = $scriptInfo["6"]
            if ($script.ContainsKey("LocalPath")) {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache -LocalPath $script.LocalPath
            } else {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
            }
        }
        "7" {
            Show-ScriptDescription "7"
            $script = $scriptInfo["7"]
            if ($script.ContainsKey("LocalPath")) {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache -LocalPath $script.LocalPath
            } else {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
            }
        }
        "8" {
            Show-ScriptDescription "8"
            $script = $scriptInfo["8"]
            if ($script.ContainsKey("LocalPath")) {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache -LocalPath $script.LocalPath
            } else {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
            }
        }
        "9" {
            Show-ScriptDescription "9"
            $script = $scriptInfo["9"]
            if ($script.ContainsKey("LocalPath")) {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache -LocalPath $script.LocalPath
            } else {
                Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
            }
        }
        
        # Configuration options
        "C" { Configure-Settings }
        "D" { Check-AllDependencies }
        "P" { Configure-PnPConnection }
        "F" { Open-ConfigFolder }
        "0" { 
            Write-Host "Exiting..." -ForegroundColor Yellow
            $exitScript = $true
        }
        default { Write-Host "❌ Invalid choice. Please try again." -ForegroundColor Red }
    }

} while (-not $exitScript)

Write-Host "Thank you for using $appName!" -ForegroundColor Green

# End of MS 365 Admin Alembic script
