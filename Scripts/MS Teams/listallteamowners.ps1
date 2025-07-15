# Since there is no way to see the owners of other Teams, this script list them all and save to CSV

Import-Module Microsoft.Graph.Users

# Get all Teams
Write-Host "Fetching all Teams..." -ForegroundColor Cyan
$Teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All:$true
$totalTeams = $Teams.Count
Write-Host "Found $totalTeams Teams." -ForegroundColor Green

# Initialize array for results
$Results = @()
$counter = 0

foreach ($Team in $Teams) {
    $counter++
    Write-Host "Processing Team $counter of $totalTeams`:` $($Team.DisplayName)" -ForegroundColor Yellow

    # Get owners of each team
    $Owners = Get-MgGroupOwner -GroupId $Team.Id
    foreach ($Owner in $Owners) {
        # Fetch full user details to ensure we get an email
        $OwnerDetails = Get-MgUser -UserId $Owner.Id -ErrorAction SilentlyContinue

        # Use Mail if available, otherwise fallback to UserPrincipalName
        $OwnerEmail = if ($OwnerDetails.Mail) { $OwnerDetails.Mail } else { $OwnerDetails.UserPrincipalName }

        $Results += [PSCustomObject]@{
            TeamName   = $Team.DisplayName
            OwnerEmail = $OwnerEmail
        }
    }
}

# Ask user for CSV output location
$saveInWorkFolder = Read-Host "Do you want to save the CSV in the current working folder? (Y/N)"
if ($saveInWorkFolder -match "^[Yy]") {
    $csvPath = "Teams_Owners.csv"
} else {
    $customPath = Read-Host "Enter the full path where you want to save the CSV file (e.g., C:\Reports\Teams_Owners.csv)"
    if ([string]::IsNullOrWhiteSpace($customPath)) {
        Write-Host "No path provided. Using current folder." -ForegroundColor Yellow
        $csvPath = "Teams_Owners.csv"
    } else {
        $csvPath = $customPath
    }
}

# Export results to CSV
$Results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "Export completed: $csvPath" -ForegroundColor Green
