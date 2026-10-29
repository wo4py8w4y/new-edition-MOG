# Import the CSV file with a column 'PrimarySMTPAddress'
$users = Import-Csv -Path ".\dls.csv"

foreach ($user in $users) {
    $mailbox = $user.PrimarySMTPAddress

    # Find all distribution groups where the user is a member
    $groups = Get-DistributionGroup | Where-Object {
        (Get-DistributionGroupMember -Identity $_.Identity -ResultSize Unlimited | Where-Object { $_.PrimarySmtpAddress -eq $mailbox })
    }

    foreach ($group in $groups) {
        # Get group owner(s)
        $owner = (Get-DistributionGroup -Identity $group.Identity).ManagedBy | Get-Recipient
        $ownerNamesStr = ($owner | Select-Object -ExpandProperty DisplayName) -join ', '

        # Prepare output array for this group
        $output = @()

        # Get group members and their permissions
        $members = Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited
        foreach ($member in $members) {
            # Get permissions for the member (if any)
            $permissions = Get-RecipientPermission -Identity $group.Identity -Trustee $member.Name -ErrorAction SilentlyContinue | 
                Select-Object -ExpandProperty AccessRights -ErrorAction SilentlyContinue

            $output += [PSCustomObject]@{
                GroupName         = $group.DisplayName
                GroupSMTP         = $group.PrimarySmtpAddress
                GroupGUID         = $group.Guid
                GroupOwner        = $ownerNamesStr
                Member            = $member.PrimarySmtpAddress
                MemberDisplayName = $member.DisplayName
                Permissions       = ($permissions -join ', ')
            }
        }

        # Export results to a separate CSV for each DL
        $safeGroupName = ($group.DisplayName -replace '[\\\/:*?"<>|]', '_')
        $output | Export-Csv -Path "$safeGroupName.csv" -NoTypeInformation
    }
}

