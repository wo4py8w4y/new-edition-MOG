#requiredModules ExchangeOnlineManagement
# This script retrieves all Distribution Groups and Shared Mailboxes from Exchange Online,
Write-Host "Make sure you have the Exchange Online Management module installed and connected." -ForegroundColor Yellow
try {
    Write-Host "Retrieving all Distribution Groups..." -ForegroundColor Cyan
    $allDGs = Get-DistributionGroup -ErrorAction Stop -ResultSize Unlimited
    $allDGs | Select-Object Name,PrimarySmtpAddress,Guid | Export-Csv -Path "DistributionGroups.csv" -NoTypeInformation
    Write-Host "Exported Distribution Groups to DistributionGroups.csv" -ForegroundColor Green
} catch {
    Write-Error "Failed to retrieve or export Distribution Groups: $_"
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

# Export all Shared Mailboxes
try {
    Write-Host "Retrieving all Shared Mailboxes..." -ForegroundColor Cyan
    $allSharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ErrorAction Stop -ResultSize Unlimited
    $allSharedMailboxes | Select-Object DisplayName,PrimarySmtpAddress,Guid | Export-Csv -Path "SharedMailboxes.csv" -NoTypeInformation
    Write-Host "Exported Shared Mailboxes to SharedMailboxes.csv" -ForegroundColor Green
} catch {
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
} catch {
    Write-Error "Failed to import Users.csv: $_"
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

# Build GUID lookup tables for names
$dgGuidToName = @{}
foreach ($dg in $allDGs) {
    $dgGuidToName[$dg.Guid.Guid] = $dg.Name
}
$mbxGuidToName = @{}
foreach ($mbx in $allSharedMailboxes) {
    $mbxGuidToName[$mbx.Guid.Guid] = $mbx.DisplayName
}

# Build Distribution Group membership lookup (by GUID)
$dgMemberships = @()
$dgMemberLookup = @{}

foreach ($dg in $allDGs) {
    try {
        Write-Host "Checking members of Distribution Group: $($dg.Name) [$($dg.Guid.Guid)]" -ForegroundColor Yellow
        $members = Get-DistributionGroupMember -Identity $dg.Guid.Guid -ResultSize Unlimited -ErrorAction Stop
        foreach ($member in $members) {
            if ($member.PrimarySmtpAddress) {
                $email = $member.PrimarySmtpAddress.ToString().ToLower()
                if (-not $dgMemberLookup.ContainsKey($email)) {
                    $dgMemberLookup[$email] = @()
                }
                $dgMemberLookup[$email] += $dg.Guid.Guid
                Write-Host "  Found member: $email in $($dg.Name) [$($dg.Guid.Guid)]" -ForegroundColor Gray
            }
        }
    } catch {
        Write-Warning "Failed to get members for Distribution Group '$($dg.Name)': $_"
        continue
    }
}

# Build Shared Mailbox permission lookup (by GUID)
$mbxMemberships = @()
$mbxPermissionLookup = @{}

foreach ($mbx in $allSharedMailboxes) {
    try {
        Write-Host "Checking permissions for Shared Mailbox: $($mbx.DisplayName) [$($mbx.Guid.Guid)]" -ForegroundColor Yellow
        $permissions = Get-MailboxPermission -Identity $mbx.Guid.Guid -ErrorAction Stop | Where-Object {
            $_.AccessRights -contains "FullAccess" -and
            $_.IsInherited -eq $false -and
            $_.User -notlike "NT AUTHORITY\SELF" -and
            $_.User -notlike "S-1-5-*" # Exclude system/security principals
        }
        foreach ($perm in $permissions) {
            $user = $perm.User.ToString().ToLower()
            if (-not $mbxPermissionLookup.ContainsKey($user)) {
                $mbxPermissionLookup[$user] = @()
            }
            $mbxPermissionLookup[$user] += $mbx.Guid.Guid
            Write-Host "  Found permission: $user has access to $($mbx.DisplayName) [$($mbx.Guid.Guid)]" -ForegroundColor Gray
        }
    } catch {
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
        foreach ($dgGuid in $dgMemberLookup[$userEmail]) {
            $dgName = $dgGuidToName[$dgGuid]
            $dgMemberships += [PSCustomObject]@{
                User              = $user.UserPrincipalName
                DistributionGroup = $dgName
                DistributionGroupGuid = $dgGuid
            }
            Write-Host "  User $userEmail is a member of Distribution Group: $dgName [$dgGuid]" -ForegroundColor Green
        }
    }

    # Shared Mailboxes
    if ($mbxPermissionLookup.ContainsKey($userEmail)) {
        foreach ($mbxGuid in $mbxPermissionLookup[$userEmail]) {
            $mbxName = $mbxGuidToName[$mbxGuid]
            $mbxMemberships += [PSCustomObject]@{
                User          = $user.UserPrincipalName
                SharedMailbox = $mbxName
                SharedMailboxGuid = $mbxGuid
            }
            Write-Host "  User $userEmail has access to Shared Mailbox: $mbxName [$mbxGuid]" -ForegroundColor Green
        }
    }
}

# Export results
try {
    $dgMemberships | Export-Csv -Path "UserDistributionGroupMemberships.csv" -NoTypeInformation
    $mbxMemberships | Export-Csv -Path "UserSharedMailboxMemberships.csv" -NoTypeInformation
    Write-Host "Exported user memberships to UserDistributionGroupMemberships.csv and UserSharedMailboxMemberships.csv" -ForegroundColor Green
} catch {
    Write-Error "Failed to export results: $_"
}

# Disconnect session

