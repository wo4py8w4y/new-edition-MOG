$mailboxes = Import-Csv -Path ".\mbx.csv"

$total = $mailboxes.Count
$counter = 0

foreach ($mb in $mailboxes) {
    $counter++
    Write-Progress -Activity "Processing mailboxes" -Status "$counter of $total" -PercentComplete (($counter / $total) * 100)

    $mbx = Get-Mailbox -Identity $mb.PrimarySMTPAddress -ErrorAction SilentlyContinue
    if ($mbx) {
        # Get owner (ManagedBy property, if any)
        $owner = $mbx.ManagedBy | Get-Recipient -ErrorAction SilentlyContinue
        $ownerName = if ($owner) { $owner.DisplayName } else { "" }

        # Get members with Full Access
        $fullAccess = Get-MailboxPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue | 
            Where-Object { $_.AccessRights -contains "FullAccess" -and $_.User -notlike "NT AUTHORITY\SELF" } |
            Select-Object -ExpandProperty User

        # Get members with Send As
        $sendAs = Get-RecipientPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue | 
            Where-Object { $_.AccessRights -contains "SendAs" } |
            Select-Object -ExpandProperty Trustee

        # Get members with Send on Behalf
        $sendOnBehalf = $mbx.GrantSendOnBehalfTo | Get-Recipient -ErrorAction SilentlyContinue | 
            Select-Object -ExpandProperty DisplayName

        # Prepare output object
        $output = [PSCustomObject]@{
            MailboxDisplayName = $mbx.DisplayName
            PrimarySMTPAddress = $mbx.PrimarySmtpAddress
            Alias              = $mbx.Alias
            GUID               = $mbx.Guid
            Owner              = $ownerName
            FullAccessMembers  = ($fullAccess -join ', ')
            SendAsMembers      = ($sendAs -join ', ')
            SendOnBehalfMembers= ($sendOnBehalf -join ', ')
        }

        # Export to a separate CSV for each mailbox
        $safeName = ($mbx.DisplayName -replace '[\\\/:*?"<>|]', '_')
        $output | Export-Csv -Path "$safeName.csv" -NoTypeInformation
    }
}

