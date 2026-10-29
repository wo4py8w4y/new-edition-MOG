$users = Import-Csv -Path .\usersnames.csv
# Get all shared mailboxes
$sharedMailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox

# Initialize an array to hold results
$results = @()
Import-Csv -Path .\usersnames.csv
foreach ($mailbox in $sharedMailboxes) {
    $permissions = Get-MailboxPermission -Identity $mailbox.Identity | Where-Object {
        $_.User -like $useremail -and $_.AccessRights -contains "FullAccess"
    }
    
    if ($permissions) {
        $results += [PSCustomObject]@{
            Mailbox      = $mailbox.DisplayName
            User         = $useremail
            AccessRights = $permissions.AccessRights -join ', '
        }
    }
}

# Output the results
$results | out-excel -path .\excelsharedmailboxes.xlsx
$results | Format-Table -AutoSize