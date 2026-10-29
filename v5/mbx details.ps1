$mailboxes = Import-Csv -Path ".\mbx.csv"

# Check if the CSV file has a column named 'DisplayName'
if (-not $mailboxes[0].PSObject.Properties.Name -contains 'DisplayName') {
    Write-Error "The CSV file must contain a column named 'DisplayName'."
    exit
}
$total = $mailboxes.Count
$counter = 0

foreach ($mb in $mailboxes) {
    $counter++
    Write-Progress -Activity "Processing mailboxes" -Status "$counter of $total" -PercentComplete (($counter / $total) * 100)

    if (-not $mb.DisplayName) {
        Write-Warning "Missing DisplayName in row $counter. Skipping."
        continue
    }

    $mbx = Get-Mailbox -Identity $mb.DisplayName -ErrorAction SilentlyContinue
    if (-not $mbx) {
        Write-Warning "Mailbox not found: $($mb.DisplayName). Skipping."
        continue
    }

    # Get owner (ManagedBy property, if any)
    $owner = $null
    $ownerName = ""
    try {
        if ($mbx.ManagedBy) {
            $owner = $mbx.ManagedBy | Get-Recipient -ErrorAction Stop
            $ownerName = ($owner | Select-Object -ExpandProperty DisplayName) -join ', '
        }
    } catch {
        Write-Warning "Could not resolve owner for mailbox $($mbx.DisplayName)."
        $ownerName = ""
    }

    # Get members with Full Access
    $fullAccess = @()
    try {
        $fullAccess = Get-MailboxPermission -Identity $mbx.Identity -ErrorAction Stop | 
            Where-Object { $_.AccessRights -contains "FullAccess" -and $_.User -notlike "NT AUTHORITY\SELF" } |
            Select-Object -ExpandProperty User
    } catch {
        Write-Warning "Could not retrieve Full Access permissions for $($mbx.DisplayName)."
    }

    # Get members with Send As
    $sendAs = @()
    try {
        $sendAs = Get-RecipientPermission -Identity $mbx.Identity -ErrorAction Stop | 
            Where-Object { $_.AccessRights -contains "SendAs" } |
            Select-Object -ExpandProperty Trustee
    } catch {
        Write-Warning "Could not retrieve Send As permissions for $($mbx.DisplayName)."
    }

    # Get members with Send on Behalf
    $sendOnBehalf = @()
    try {
        if ($mbx.GrantSendOnBehalfTo) {
            $sendOnBehalf = $mbx.GrantSendOnBehalfTo | Get-Recipient -ErrorAction Stop | 
                Select-Object -ExpandProperty DisplayName
        }
    } catch {
        Write-Warning "Could not retrieve Send on Behalf permissions for $($mbx.DisplayName)."
    }

    # Prepare output object
    $output = [PSCustomObject]@{
        MailboxDisplayName   = $mbx.DisplayName
        PrimarySMTPAddress   = $mbx.PrimarySmtpAddress
        Alias                = $mbx.Alias
        GUID                 = $mbx.Guid
        Owner                = $ownerName
        FullAccessMembers    = ($fullAccess -join ', ')
        SendAsMembers        = ($sendAs -join ', ')
        SendOnBehalfMembers  = ($sendOnBehalf -join ', ')
    }

    # Export to a separate CSV for each mailbox
    $safeName = ($mbx.DisplayName -replace '[\\\/:*?"<>|]', '_')
    try {
        $output | Export-Csv -Path "$safeName.csv" -NoTypeInformation
    } catch {
        Write-Warning "Failed to export CSV for $($mbx.DisplayName)."
    }
}

