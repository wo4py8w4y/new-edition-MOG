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
    $allDGs = Get-DistributionGroup -ErrorAction Stop -ResultSize Unlimited
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
    $allSharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ErrorAction Stop -ResultSize Unlimited
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

foreach ($dg in $allDGs) {
    try {
        Write-Host "Checking members of Distribution Group: $($dg.Name)" -ForegroundColor Yellow
        $members = Get-DistributionGroupMember -Identity $dg.Identity -ResultSize Unlimited -ErrorAction Stop
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
        $permissions = Get-MailboxPermission -Identity $mbx.Identity -ErrorAction Stop | Where-Object {
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
            $mbxPermissionLookup[$user] += $mbx.DisplayName
            Write-Host "  Found permission: $user has access to $($mbx.DisplayName)" -ForegroundColor Gray
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
        foreach ($mbx in $mbxPermissionLookup[$userEmail]) {
            $mbxMemberships += [PSCustomObject]@{
                User          = $user.UserPrincipalName
                SharedMailbox = $mbx
            }
            Write-Host "  User $userEmail has access to Shared Mailbox: $mbx" -ForegroundColor Green
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
