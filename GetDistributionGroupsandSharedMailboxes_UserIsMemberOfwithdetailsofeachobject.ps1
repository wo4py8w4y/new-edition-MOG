param(
    [ValidateRange(1, 1000)]
    [int]$BatchSize = 25,

    [ValidateRange(0, 3600)]
    [int]$BatchDelaySeconds = 10,

    [ValidateRange(0, 60)]
    [int]$RequestIntervalSeconds = 1,

    [ValidateRange(0, 10)]
    [int]$MaxRetryCount = 5
)

# Changelog
# 2026-10-29 - Added chunk-based pacing and transient throttling retries for Exchange Online requests.

function Invoke-ExchangeOnlineRequest {
    param(
        [Parameter(Mandatory)]
        [string]$OperationName,

        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock
    )

    $attempt = 0
    while ($true) {
        if ($script:ExchangeRequestCount -gt 0) {
            $waitUntil = $script:LastExchangeRequestTime.AddSeconds($RequestIntervalSeconds)
            $waitMilliseconds = [int][math]::Ceiling(($waitUntil - [datetime]::UtcNow).TotalMilliseconds)
            if ($waitMilliseconds -gt 0) {
                Start-Sleep -Milliseconds $waitMilliseconds
            }

            if (($script:ExchangeRequestCount % $BatchSize) -eq 0 -and $BatchDelaySeconds -gt 0) {
                Write-Verbose "Completed $($script:ExchangeRequestCount) Exchange Online requests; pausing for $BatchDelaySeconds seconds."
                Start-Sleep -Seconds $BatchDelaySeconds
            }
        }

        $script:ExchangeRequestCount++
        $script:LastExchangeRequestTime = [datetime]::UtcNow
        $attempt++

        try {
            Write-Verbose "Exchange Online request $($script:ExchangeRequestCount): $OperationName (attempt $attempt)."
            & $ScriptBlock
            return
        }
        catch {
            $isTransient = $_.Exception -is [System.TimeoutException] -or
                $_.Exception.Message -match '(?i)throttl|too many requests|server busy|temporarily unavailable|timeout|transient|\b429\b|\b503\b'

            if (-not $isTransient -or $attempt -gt $MaxRetryCount) {
                throw
            }

            $retryDelaySeconds = [math]::Min(60, [math]::Pow(2, $attempt))
            Write-Warning "$OperationName failed with a transient Exchange Online error. Retrying in $retryDelaySeconds seconds (attempt $attempt of $MaxRetryCount): $($_.Exception.Message)"
            Start-Sleep -Seconds $retryDelaySeconds
        }
    }
}

$script:ExchangeRequestCount = 0
$script:LastExchangeRequestTime = [datetime]::MinValue

<#Connect to Exchange Online (requires Exchange Online PowerShell module)
try {
    Connect-ExchangeOnline -ErrorAction Stop
} catch {
    Write-Error "Failed to connect to Exchange Online: $_"
    exit 1
}
#>
# Export all Distribution Groups

# Export all Distribution Groups
try {
    Write-Host "Retrieving all Distribution Groups..." -ForegroundColor Cyan
    $allDGs = Invoke-ExchangeOnlineRequest -OperationName 'Get-DistributionGroup inventory' -ScriptBlock {
        Get-DistributionGroup -ErrorAction Stop -ResultSize Unlimited
    }
    $allDGs | Select-Object Name, PrimarySmtpAddress | Export-Csv -Path "DistributionGroups.csv" -NoTypeInformation
    Write-Host "Exported Distribution Groups to DistributionGroups.csv" -ForegroundColor Green
}
catch {
    Write-Error "Failed to retrieve or export Distribution Groups: $_"
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

# Export all Shared Mailboxes
try {
    Write-Host "Retrieving all Shared Mailboxes..." -ForegroundColor Cyan
    $allSharedMailboxes = Invoke-ExchangeOnlineRequest -OperationName 'Get-Mailbox shared mailbox inventory' -ScriptBlock {
        Get-Mailbox -RecipientTypeDetails SharedMailbox -ErrorAction Stop -ResultSize Unlimited
    }
    $allSharedMailboxes | Select-Object DisplayName, PrimarySmtpAddress | Export-Csv -Path "SharedMailboxes.csv" -NoTypeInformation
    Write-Host "Exported Shared Mailboxes to SharedMailboxes.csv" -ForegroundColor Green
}
catch {
    Write-Error "Failed to retrieve or export Shared Mailboxes: $_"
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

# Import users from CSV (assume column 'UserPrincipalName' or 'Email')
try {
    Write-Host "Importing users from Users.csv..." -ForegroundColor Cyan
    $users = Import-Csv -Path "Users.csv" -ErrorAction Stop
    if (-not $users -or -not $users[0].PSObject.Properties.Match('UserPrincipalName')) {
        Write-Error "Users.csv must contain a 'UserPrincipalName' column."
        Disconnect-ExchangeOnline -Confirm:$false
        exit 1
    }
    $userEmails = $users | ForEach-Object { $_.UserPrincipalName.ToLower() }
    Write-Host "Imported $($users.Count) users." -ForegroundColor Green
}
catch {
    Write-Error "Failed to import Users.csv: $_"
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

# Build Distribution Group membership lookup
$dgMemberships = @()
$dgMemberLookup = @{}
$dgMemberCache = @{}

foreach ($dg in $allDGs) {
    try {
        Write-Host "Checking members of Distribution Group: $($dg.Name)" -ForegroundColor Yellow
        $members = Invoke-ExchangeOnlineRequest -OperationName "Get-DistributionGroupMember '$($dg.Name)'" -ScriptBlock {
            Get-DistributionGroupMember -Identity $dg.Identity -ResultSize Unlimited -ErrorAction Stop
        }
        $dgMemberCache[$dg.Name] = @($members)
        foreach ($member in $members) {
            if ($member.PrimarySmtpAddress) {
                $email = $member.PrimarySmtpAddress.ToString().ToLower()
                if (-not $dgMemberLookup.ContainsKey($email)) {
                    $dgMemberLookup[$email] = @()
                }
                $dgMemberLookup[$email] += $dg.Name
                Write-Host "  Found member: $email in $($dg.Name)" -ForegroundColor Gray
            }
        }
    }
    catch {
        Write-Warning "Failed to get members for Distribution Group '$($dg.Name)': $_"
        continue
    }
}

# Build Shared Mailbox permission lookup
$mbxMemberships = @()
$mbxPermissionLookup = @{}

foreach ($mbx in $allSharedMailboxes) {
    try {
        Write-Host "Checking permissions for Shared Mailbox: $($mbx.DisplayName)" -ForegroundColor Yellow

        # FullAccess
        $permissions = Invoke-ExchangeOnlineRequest -OperationName "Get-MailboxPermission '$($mbx.DisplayName)'" -ScriptBlock {
            Get-MailboxPermission -Identity $mbx.Identity -ErrorAction Stop
        } | Where-Object {
            $_.AccessRights -contains "FullAccess" -and
            $_.IsInherited -eq $false -and
            $_.User -notlike "NT AUTHORITY\SELF" -and
            $_.User -notlike "S-1-5-*"
        }
        foreach ($perm in $permissions) {
            $user = $perm.User.ToString().ToLower()
            if (-not $mbxPermissionLookup.ContainsKey($user)) {
                $mbxPermissionLookup[$user] = @()
            }
            $mbxPermissionLookup[$user] += [PSCustomObject]@{
                SharedMailbox = $mbx.DisplayName
                AccessType    = "FullAccess"
            }
            Write-Host "  Found FullAccess: $user has access to $($mbx.DisplayName)" -ForegroundColor Gray
        }

        # SendAs
        try {
            $sendAsPerms = Invoke-ExchangeOnlineRequest -OperationName "Get-RecipientPermission '$($mbx.DisplayName)'" -ScriptBlock {
                Get-RecipientPermission -Identity $mbx.Identity -ErrorAction Stop
            } | Where-Object {
                $_.Trustee -notlike "NT AUTHORITY\SELF" -and
                $_.Trustee -notlike "S-1-5-*"
            }
        }
        catch {
            Write-Warning "Failed to get SendAs permissions for Shared Mailbox '$($mbx.DisplayName)': $_"
            $sendAsPerms = @()
        }
        foreach ($perm in $sendAsPerms) {
            $user = $perm.Trustee.ToString().ToLower()
            if (-not $mbxPermissionLookup.ContainsKey($user)) {
                $mbxPermissionLookup[$user] = @()
            }
            $mbxPermissionLookup[$user] += [PSCustomObject]@{
                SharedMailbox = $mbx.DisplayName
                AccessType    = "SendAs"
            }
            Write-Host "  Found SendAs: $user has SendAs on $($mbx.DisplayName)" -ForegroundColor Gray
        }

        # Send on Behalf
        try {
            $mbxProps = Invoke-ExchangeOnlineRequest -OperationName "Get-Mailbox delegate properties '$($mbx.DisplayName)'" -ScriptBlock {
                Get-Mailbox -Identity $mbx.Identity -ErrorAction Stop
            }
        }
        catch {
            Write-Warning "Failed to get Send on Behalf delegates for Shared Mailbox '$($mbx.DisplayName)': $_"
            $mbxProps = $null
        }
        if ($mbxProps.GrantSendOnBehalfTo) {
            foreach ($delegate in $mbxProps.GrantSendOnBehalfTo) {
                $user = $delegate.Name.ToLower()
                if (-not $mbxPermissionLookup.ContainsKey($user)) {
                    $mbxPermissionLookup[$user] = @()
                }
                $mbxPermissionLookup[$user] += [PSCustomObject]@{
                    SharedMailbox = $mbx.DisplayName
                    AccessType    = "SendOnBehalf"
                }
                Write-Host "  Found SendOnBehalf: $user has SendOnBehalf on $($mbx.DisplayName)" -ForegroundColor Gray
            }
        }
    }
    catch {
        Write-Warning "Failed to get permissions for Shared Mailbox '$($mbx.DisplayName)': $_"
        continue
    }
}

# Match users to Distribution Groups and Shared Mailboxes
foreach ($user in $users) {
    $userEmail = $user.UserPrincipalName.ToLower()
    Write-Host "Processing user: $userEmail" -ForegroundColor Cyan

    # Distribution Groups
    if ($dgMemberLookup.ContainsKey($userEmail)) {
        foreach ($dg in $dgMemberLookup[$userEmail]) {
            $dgMemberships += [PSCustomObject]@{
                User              = $user.UserPrincipalName
                DistributionGroup = $dg
            }
            Write-Host "  User $userEmail is a member of Distribution Group: $dg" -ForegroundColor Green
        }
    }

    # Shared Mailboxes
    if ($mbxPermissionLookup.ContainsKey($userEmail)) {
        foreach ($perm in $mbxPermissionLookup[$userEmail]) {
            $mbxMemberships += [PSCustomObject]@{
                User          = $user.UserPrincipalName
                SharedMailbox = $perm.SharedMailbox
                AccessType    = $perm.AccessType
            }
            Write-Host "  User $userEmail has $($perm.AccessType) on Shared Mailbox: $($perm.SharedMailbox)" -ForegroundColor Green
        }
    }
}

# Export results
try {
    $dgMemberships | Export-Csv -Path "UserDistributionGroupMemberships.csv" -NoTypeInformation
    $mbxMemberships | Export-Csv -Path "UserSharedMailboxMemberships.csv" -NoTypeInformation
    Write-Host "Exported user memberships to UserDistributionGroupMemberships.csv and UserSharedMailboxMemberships.csv" -ForegroundColor Green
}
catch {
    Write-Error "Failed to export results: $_"
}

# --- Filtered Shared Mailboxes and Distribution Groups for Reports ---
$usedSharedMailboxes = $mbxMemberships | Select-Object -ExpandProperty SharedMailbox -Unique
$usedDistributionGroups = $dgMemberships | Select-Object -ExpandProperty DistributionGroup -Unique

# --- Shared Mailbox Summary Report ---
$sharedMailboxSummary = foreach ($mbx in $allSharedMailboxes | Where-Object { $usedSharedMailboxes -contains $_.DisplayName }) {
    try {
        $mbxStats = Invoke-ExchangeOnlineRequest -OperationName "Get-MailboxStatistics '$($mbx.DisplayName)'" -ScriptBlock {
            Get-MailboxStatistics -Identity $mbx.Identity -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "Failed to get statistics for Shared Mailbox '$($mbx.DisplayName)': $_"
        $mbxStats = $null
    }
    $owner = if ($mbx.ManagedBy) { ($mbx.ManagedBy | ForEach-Object { $_.ToString() }) -join '; ' } else { "" }
    [PSCustomObject]@{
        DisplayName      = $mbx.DisplayName
        PrimarySMTP      = $mbx.PrimarySmtpAddress
        ProxyAddresses   = ($mbx.EmailAddresses | Where-Object { $_ -like "smtp:*" }) -join '; '
        Owner            = $owner
        SizeMB           = if ($mbxStats -and $mbxStats.TotalItemSize) {
                               if ($mbxStats.TotalItemSize.ToString() -match "([\d\.]+) MB") {
                                   [math]::Round([double]$matches[1],2)
                               } else { "" }
                           } else { "" }
        ItemCount        = if ($mbxStats) { $mbxStats.ItemCount } else { "" }
        WhenCreated      = $mbx.WhenCreated
        HiddenFromGAL    = $mbx.HiddenFromAddressListsEnabled
        OrganizationalUnit = $mbx.OrganizationalUnit
    }
}
$sharedMailboxSummary | Export-Csv -Path "SharedMailboxSummary.csv" -NoTypeInformation

# --- Per Shared Mailbox Membership Report ---
foreach ($mbxName in $usedSharedMailboxes) {
    $safeName = $mbxName -replace '[\\/:*?"<>|]', '_'
    $members = $mbxMemberships | Where-Object { $_.SharedMailbox -eq $mbxName }
    $members | Export-Csv -Path "SharedMailboxMembership-$safeName.csv" -NoTypeInformation
}

# --- Distribution List Summary Report ---
$dgSummary = foreach ($dg in $allDGs | Where-Object { $usedDistributionGroups -contains $_.Name }) {
    $owner = ($dg.ManagedBy | ForEach-Object { $_.ToString() }) -join '; '
    [PSCustomObject]@{
        Name            = $dg.Name
        PrimarySMTP     = $dg.PrimarySmtpAddress
        Alias           = $dg.Alias
        ProxyAddresses  = ($dg.EmailAddresses | Where-Object { $_ -like "smtp:*" }) -join '; '
        Owner           = $owner
        HiddenFromGAL   = $dg.HiddenFromAddressListsEnabled
        Type            = $dg.RecipientTypeDetails
        WhenCreated     = $dg.WhenCreated
    }
}
$dgSummary | Export-Csv -Path "DistributionGroupSummary.csv" -NoTypeInformation

# --- Per Distribution List Membership Report ---
foreach ($dgName in $usedDistributionGroups) {
    $safeName = $dgName -replace '[\\/:*?"<>|]', '_'
    $members = @()
    try {
        $dg = $allDGs | Where-Object { $_.Name -eq $dgName }
        if ($dgMemberCache.ContainsKey($dgName)) {
            $dgMembers = $dgMemberCache[$dgName]
        }
        else {
            $dgMembers = Invoke-ExchangeOnlineRequest -OperationName "Get-DistributionGroupMember report '$dgName'" -ScriptBlock {
                Get-DistributionGroupMember -Identity $dg.Identity -ResultSize Unlimited -ErrorAction Stop
            }
        }
        foreach ($member in $dgMembers) {
            $members += [PSCustomObject]@{
                DistributionGroup = $dgName
                MemberName        = $member.DisplayName
                MemberType        = $member.RecipientType
                MemberEmail       = $member.PrimarySmtpAddress
            }
        }
    } catch {
        $members += [PSCustomObject]@{
            DistributionGroup = $dgName
            MemberName        = "Error retrieving members"
            MemberType        = ""
            MemberEmail       = ""
        }
    }
    $members | Export-Csv -Path "DistributionGroupMembership-$safeName.csv" -NoTypeInformation
}
