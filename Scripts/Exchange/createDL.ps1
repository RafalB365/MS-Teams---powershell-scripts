# This spript is creating an exchange Distribution List based on the MS 365 group members 

# Connect to Exchange Online (Interactive sign-in)
Connect-ExchangeOnline

# Set your group and DL names
$SourceGroup = "XXX"         # <-- CHANGE THIS
$TargetDL    = "XXX"             # <-- CHANGE THIS

# Get members of the source M365 Group
Write-Host "Getting members of the M365 group: $SourceGroup"
$GroupMembers = Get-UnifiedGroupLinks -Identity $SourceGroup -LinkType Members

# Get existing members of the target DL
Write-Host "Getting existing members of DL: $TargetDL"
$DL_Members = Get-DistributionGroupMember -Identity $TargetDL

# Collect the email addresses of existing DL members, normalized to lower case for comparison
$DL_Emails = $DL_Members | ForEach-Object {
    if ($_.PrimarySmtpAddress)      { $_.PrimarySmtpAddress.ToString().ToLower() }
    elseif ($_.EmailAddress)        { $_.EmailAddress.ToString().ToLower() }
    elseif ($_.ExternalEmailAddress){ $_.ExternalEmailAddress.ToString().ToLower() }
    elseif ($_.WindowsEmailAddress) { $_.WindowsEmailAddress.ToString().ToLower() }
    else { $null }
} | Where-Object { $_ }

$Idx = 1
$Total = $GroupMembers.Count

foreach ($Member in $GroupMembers) {

    # Try common property names for email address
    $Email = $null
    if ($Member.PrimarySmtpAddress)       { $Email = $Member.PrimarySmtpAddress.ToString().ToLower() }
    elseif ($Member.EmailAddress)         { $Email = $Member.EmailAddress.ToString().ToLower() }
    elseif ($Member.ExternalEmailAddress) { $Email = $Member.ExternalEmailAddress.ToString().ToLower() }
    elseif ($Member.WindowsEmailAddress)  { $Email = $Member.WindowsEmailAddress.ToString().ToLower() }

    if (![string]::IsNullOrWhiteSpace($Email)) {
        if ($DL_Emails -notcontains $Email) {
            Write-Host "[$Idx of $Total] Adding missing member $Email to $TargetDL"
            try {
                Add-DistributionGroupMember -Identity $TargetDL -Member $Email -ErrorAction Stop
                Write-Host "   -> Added successfully" -ForegroundColor Green
            } catch {
                $errMsg = $_.Exception.Message
                Write-Host "   -> Failed to add $Email" $errMsg -ForegroundColor Red
            }
        } else {
            Write-Host "[$Idx of $Total] Member $Email is already in $TargetDL, skipping" -ForegroundColor Cyan
        }
    } else {
        Write-Host "[$Idx of $Total] Skipped: No valid email address for member: $($Member.Name)" -ForegroundColor Yellow
    }
    $Idx++
}

Write-Host "All done!" -ForegroundColor Magenta