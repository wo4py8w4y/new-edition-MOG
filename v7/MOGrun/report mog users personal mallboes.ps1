# This script will pull a list of emails (or upnames) from a CSV file, 
#look up their mailbox details using the Microsoft Graph API, and then check for shared mailbox permissions. It will output the results to the console and save them to a CSV file.
Connect-ExchangeOnline
Connect-MgGraph -NoWelcome -Scopes Reports.Read.All

# Ensure Graph returns non-obfuscated user data
$ObfuscatedReset = $False
If ((Get-MgBetaAdminReportSetting).DisplayConcealedNames -eq $True) {
    $Parameters = @{ displayConcealedNames = $False }
    Update-MgBetaAdminReportSetting -BodyParameter $Parameters
    $ObfuscatedReset = $True
}

Write-Host "Fetching Microsoft 365 usage data"
Get-MgReportEmailActivityUserDetail -Period 'D180' -Outfile EmailActivity.CSV 
[array]$EmailActivityData = Import-CSV EmailActivity.CSV
Get-MgReportMailboxUsageDetail -Period 'D180' -Outfile MailboxUsage.CSV
[array]$MailboxUsage = Import-CSV MailboxUsage.CSV

# Set log file path in current directory
$LogFile = Join-Path -Path (Get-Location) -ChildPath "MailboxProcessing.log"

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    Write-Host $logMessage
    Add-Content -Path $LogFile -Value $logMessage
}

# Import user list from CSV
$userList = Import-Csv -Path ".\users.csv" -ErrorAction Stop
Write-Host -ForegroundColor Green "User list imported successfully"
Write-Host "Total users to process: $($userList.Count)"

$userMailboxesArray = @()
$currentuser = 0

foreach ($user in $userList) {
    try {
        $currentuser++
        Write-Progress -Activity "Looking up mailboxes of user list" -Status "Mailbox $currentuser of $($userList.Count)" -PercentComplete (($currentuser / $userList.Count) * 100)
        $mailbox = Get-ExoMailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
        if ($mailbox) {
            $userMailboxesArray += $mailbox
        }
        else {
            Write-Log ("Mailbox not found for {0}" -f $user.UserPrincipalName)
        }
    }
    catch {
        Write-Log ("Error getting mailbox for {0}: {1}" -f $user.UserPrincipalName, $_)
    }
}

[array]$Mbx = $userMailboxesArray
Write-Log ("Processing {0} mailboxes..." -f $Mbx.Count)
$OutputReport = [System.Collections.Generic.List[Object]]::new()

$currentMailbox = 0
foreach ($M in $Mbx) {
    $currentMailbox++
    Write-Log ("Processing mailbox: {0} ({1}) [{2}/{3}]" -f $M.DisplayName, $M.UserPrincipalName, $currentMailbox, $Mbx.Count)
    $UserStats = $EmailActivityData | Where-Object { $_.'User Principal Name' -eq $M.UserPrincipalName }
    $MailboxStats = $MailboxUsage | Where-Object { $_.'User Principal Name' -eq $M.UserPrincipalName }

    $PrimarySMTP = $M.PrimarySmtpAddress
    $ProxyAddresses = ($M.ProxyAddresses -join "; ")
    $Alias = $M.Alias
    $Guid = $M.ExchangeGuid
    $MailboxLocation = $M.ServerName
    $ArchiveLocation = $M.ArchiveDatabase
    $RecipientType = $M.RecipientTypeDetails
    $MailboxType = $M.MailboxTypeDetail
    $ArchiveStatus = $M.ArchiveStatus
    $RetentionPolicy = $M.RetentionPolicy
    $LitHoldEnabled = $M.LitigationHoldEnabled
    $SingleItemRecovery = $M.SingleItemRecoveryEnabled
    $AutoExpandingArchive = $M.AutoExpandingArchiveEnabled

    If ($MailboxStats -and -not [string]::IsNullOrWhiteSpace($MailboxStats.'Last Activity Date')) {
        try {
            $parsedDate = Get-Date $MailboxStats.'Last Activity Date'
            $LastActivityDate = $parsedDate.ToString('dd-MM-yyyy HH:mm:ss')
            $DaysSinceActivity = (New-TimeSpan -Start $parsedDate -End (Get-Date)).Days
            Write-Log ("Last Activity Date for {0}: {1}" -f $M.UserPrincipalName, $LastActivityDate)
        }
        catch {
            $LastActivityDate = $MailboxStats.'Last Activity Date'
            $DaysSinceActivity = $Null
            Write-Log ("Error parsing Last Activity Date for {0}: {1}" -f $M.UserPrincipalName, $MailboxStats.'Last Activity Date')
        }
    }
    else {
        $LastActivityDate = $Null
        $DaysSinceActivity = $Null
        Write-Log ("No Last Activity Date found for {0}" -f $M.UserPrincipalName)
    }

    [long]$MbxStorage = $MailboxStats.'Storage Used (Byte)'
    If ($MbxStorage -le 1GB) {
        $MailboxSize = [Math]::Round(($MbxStorage / 1MB), 2)
        [string]$MbxSize = $MailboxSize.ToString() + " MB"
    }
    else {
        $MailboxSize = [Math]::Round(($MbxStorage / 1GB), 2)
        [string]$MbxSize = $MailboxSize.ToString() + " GB"
    }
    Write-Log ("Mailbox size for {0}: {1}" -f $M.UserPrincipalName, $MbxSize)

    # Archive mailbox details
    $ArchiveStats = $null
    try {
        $ArchiveStats = Get-ExoMailboxStatistics -Identity $M.UserPrincipalName -Archive -ErrorAction SilentlyContinue
    }
    catch {}
    if ($ArchiveStats) {
        [long]$ArchiveStorage = $ArchiveStats.TotalItemSize.Value.ToBytes()
        if ($ArchiveStorage -le 1GB) {
            $ArchiveSize = [Math]::Round(($ArchiveStorage / 1MB), 2).ToString() + " MB"
        }
        else {
            $ArchiveSize = [Math]::Round(($ArchiveStorage / 1GB), 2).ToString() + " GB"
        }
        $ArchiveItemCount = $ArchiveStats.ItemCount

        # Check archive single item recovery and auto-expanding status
        $ArchiveSingleItemRecovery = if ($ArchiveStats.SingleItemRecoveryEnabled) { "Enabled" } else { "Disabled" }
        $ArchiveAutoExpanding = if ($ArchiveStats.AutoExpandingArchiveEnabled) { "Enabled" } else { "Disabled" }
    }
    else {
        $ArchiveSize = "N/A"
        $ArchiveItemCount = "N/A"
        $ArchiveSingleItemRecovery = "N/A"
        $ArchiveAutoExpanding = "N/A"
    }
    Write-Log ("Archive size for {0}: {1}" -f $M.UserPrincipalName, $ArchiveSize)
    Write-Log ("Archive single item recovery for {0}: {1}" -f $M.UserPrincipalName, $ArchiveSingleItemRecovery)
    Write-Log ("Archive auto-expanding for {0}: {1}" -f $M.UserPrincipalName, $ArchiveAutoExpanding)

    $ReportLine = [PSCustomObject]@{
        UPN                       = $M.UserPrincipalName
        Name                      = $M.DisplayName
        Alias                     = $Alias
        PrimarySMTPAddress        = $PrimarySMTP
        ProxyAddresses            = $ProxyAddresses
        GUID                      = $Guid
        MailboxLocation           = $MailboxLocation
        ArchiveLocation           = $ArchiveLocation
        RecipientType             = $RecipientType
        MailboxType               = $MailboxType
        ArchiveStatus             = $ArchiveStatus
        RetentionPolicy           = $RetentionPolicy
        LitigationHoldEnabled     = $LitHoldEnabled
        SingleItemRecovery        = $SingleItemRecovery
        AutoExpandingArchive      = $AutoExpandingArchive
        Items                     = $MailboxStats.'ItemCount'
        Size                      = $MbxSize
        ArchiveSize               = $ArchiveSize
        ArchiveItemCount          = $ArchiveItemCount
        ArchiveSingleItemRecovery = $ArchiveSingleItemRecovery
        ArchiveAutoExpanding      = $ArchiveAutoExpanding
        LastActivity              = $LastActivityDate
        DaysSinceActivity         = $DaysSinceActivity
        'Send Count'              = $UserStats.'Send Count'
        'Receive Count'           = $UserStats.'Receive Count'
    }
    $OutputReport.Add($ReportLine)
}

# Output to CSV and Excel
$csvPath = Join-Path -Path (Get-Location) -ChildPath "MailboxStatisticsReport.csv"
$xlsxPath = Join-Path -Path (Get-Location) -ChildPath "MailboxStatisticsReport.xlsx"

$OutputReport | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

if (Get-Module -ListAvailable -Name ImportExcel) {
    $OutputReport | Export-Excel -Path $xlsxPath -AutoSize -WorksheetName 'MailboxStats'
    Write-Log ("Excel report saved to {0}" -f $xlsxPath)
}
else {
    Write-Log "ImportExcel module not found. Excel report not created."
}

Write-Log ("CSV report saved to {0}" -f $csvPath)

# Reset obfuscated user data setting if necessary
If ($ObfuscatedReset -eq $True) {
    $Parameters = @{ displayConcealedNames = $True }
    Update-MgBetaAdminReportSetting -BodyParameter $Parameters
}

# This is the end of the script... dont hang around here
Write-Log "Script completed successfully."
# 