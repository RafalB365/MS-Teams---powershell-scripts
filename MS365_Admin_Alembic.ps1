# MS 365 Admin Alembic with Encrypted Storage and Module Management
# Version 1.0.5
$toolkitVersion = "1.0.5"

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
Write-Host "  ███╗   ███╗██████╗ ██████╗  ██████╗ ███████╗" -ForegroundColor DarkGreen
Write-Host "  ████╗ ████║╚════██╗╚════██╗██╔════╝ ██╔════╝" -ForegroundColor DarkGreen
Write-Host "  ██╔████╔██║ █████╔╝ █████╔╝███████╗ ███████╗" -ForegroundColor DarkGreen
Write-Host "  ██║╚██╔╝██║ ╚═══██╗ ╚═══██╗╚════██║ ╚════██║" -ForegroundColor DarkGreen
Write-Host "  ██║ ╚═╝ ██║██████╔╝██████╔╝███████║ ███████║" -ForegroundColor DarkGreen
Write-Host "  ╚═╝     ╚═╝╚═════╝ ╚═════╝ ╚══════╝ ╚══════╝" -ForegroundColor DarkGreen
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
Write-Host "   ║                    v$toolkitVersion - PowerShell Toolkit                    ║" -ForegroundColor DarkGreen
Write-Host "   ║              Your MS365 Administration Swiss Army Knife!              ║" -ForegroundColor DarkGreen
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
        Name = "SharePoint Site Usage Report"
        Description = "Generates a report of site usage across your SharePoint tenant"
        Url = "https://raw.githubusercontent.com/RafalB365/MS-Teams---powershell-scripts/refs/heads/main/createDL.ps1"  # Replace with actual script
        ScriptName = "SPSiteUsage.ps1"
        RequiredModules = @("PnP.PowerShell", "Microsoft.Graph.Sites")
        RequiredScopes = @("Sites.Read.All")
        Category = "SharePoint"
    }
    
    # MS Teams scripts
    "2" = @{
        Name = "Save all Teams owners to CSV"
        Description = "Since there is no way to see the owners of other Teams, this script list them all and save to CSV"
        Url = "https://raw.githubusercontent.com/RafalB365/MS-Teams---powershell-scripts/refs/heads/main/listallteamowners.ps1"
        ScriptName = "listallteamowners.ps1"
        RequiredModules = @("Microsoft.Graph.Users", "MicrosoftTeams")
        RequiredScopes = @("GroupMember.Read.All", "Group.Read.All", "User.Read.All")
        Category = "MS Teams"
    }
    "3" = @{
        Name = "Teams Usage Analytics" 
        Description = "Generates detailed analytics about Teams usage across your organization"
        Url = "https://raw.githubusercontent.com/RafalB365/MS-Teams---powershell-scripts/refs/heads/main/createDL.ps1"  # Replace with actual script
        ScriptName = "TeamsUsageAnalytics.ps1"
        RequiredModules = @("MicrosoftTeams", "Microsoft.Graph.Reports")
        RequiredScopes = @("Reports.Read.All")
        Category = "MS Teams"
    }
    
    # Exchange scripts
    "4" = @{
        Name = "Create DL based on MS 365 group"
        Description = "This script is creating an exchange Distribution List based on the MS 365 group members"
        Url = "https://raw.githubusercontent.com/RafalB365/MS-Teams---powershell-scripts/refs/heads/main/createDL.ps1"
        ScriptName = "createDL.ps1"
        RequiredModules = @("Microsoft.Graph.Groups", "ExchangeOnlineManagement")
        RequiredScopes = @("GroupMember.ReadWrite.All", "Group.Read.All")
        Category = "Exchange"
    }
    
    # Entra (Azure AD) scripts
    "5" = @{
        Name = "User License Report"
        Description = "Generates a detailed report of license assignments across your tenant"
        Url = "https://raw.githubusercontent.com/RafalB365/MS-Teams---powershell-scripts/refs/heads/main/createDL.ps1"  # Replace with actual script
        ScriptName = "UserLicenseReport.ps1"
        RequiredModules = @("Microsoft.Graph.Users", "Microsoft.Graph.Identity.DirectoryManagement")
        RequiredScopes = @("User.Read.All", "Directory.Read.All")
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
    
    # Get unique categories
    $categories = $scriptInfo.Values | ForEach-Object { $_.Category } | Select-Object -Unique | Sort-Object
    
    foreach ($category in $categories) {
        Write-Host "`n[$category]" -ForegroundColor Cyan
        $categoryScripts = $scriptInfo.GetEnumerator() | Where-Object { $_.Value.Category -eq $category } | Sort-Object Name
        
        foreach ($script in $categoryScripts) {
            Write-Host "$($script.Key). $($script.Value.Name)"
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

# Function to execute script from GitHub
function Execute-ScriptFromGitHub {
    param (
        [string]$ScriptUrl,
        [string]$ScriptName,
        [array]$RequiredModules,
        [array]$RequiredScopes,
        [bool]$UseCache = $true
    )
    
    try {
        Write-Host "Downloading and executing $ScriptName..." -ForegroundColor Yellow
        
        # Download script content
        $scriptContent = Invoke-RestMethod -Uri $ScriptUrl -ErrorAction Stop
        
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
    
    foreach ($module in $allModules) {
        Write-Host "Checking $module..." -NoNewline
        
        try {
            $installedModule = Get-Module -Name $module -ListAvailable -ErrorAction SilentlyContinue
            if ($installedModule) {
                Write-Host " ✓ Installed (v$($installedModule.Version))" -ForegroundColor Green
            } else {
                Write-Host " ✗ Not installed" -ForegroundColor Red
                $installChoice = Read-Host "Install $module? (y/n)"
                if ($installChoice -eq 'y' -or $installChoice -eq 'Y') {
                    Install-Module -Name $module -Scope CurrentUser -Force
                    Write-Host "$module installed successfully." -ForegroundColor Green
                }
            }
        }
        catch {
            Write-Host " ✗ Error checking module: $_" -ForegroundColor Red
        }
    }
    
    Write-Host "`nDependency check completed." -ForegroundColor Green
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
            Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
        }
        "2" {
            Show-ScriptDescription "2"
            $script = $scriptInfo["2"]
            Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
        }
        "3" {
            Show-ScriptDescription "3"
            $script = $scriptInfo["3"]
            Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
        }
        "4" {
            Show-ScriptDescription "4"
            $script = $scriptInfo["4"]
            Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
        }
        "5" {
            Show-ScriptDescription "5"
            $script = $scriptInfo["5"]
            Execute-ScriptFromGitHub -ScriptUrl $script.Url -ScriptName $script.ScriptName -RequiredModules $script.RequiredModules -RequiredScopes $script.RequiredScopes -UseCache:$global:Settings.UseCache
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
