# Step 2: Import users from CSV


# Step 3: Get all shared mailboxes
# $sharedMailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox
    $sharedMailboxes = Import-Csv -Path '.\QSA-SharedMailboxes.txt'
    $users =  Import-Csv -Path ".\Usermaiboxes.csv"
# Step 4: Initialize an array to store results
$results = @()
$count = 0
# Step 5: Check permissions for each user in the CSV file
foreach ($user in $users) {
	$count + 1
    $currentUser = $user.PrimarySMTPAddress

    foreach ($mailbox in $sharedMailboxes) {
        $permissions = Get-MailboxPermission -Identity $mailbox.Identity

        # Check if the current user has access
        $userPermission = $permissions | Where-Object { $_.User -eq $currentUser }

        if ($userPermission) {
            $output = "$($mailbox.DisplayName): $currentUser - Permissions: $($userPermission.AccessRights)"
         if($count = 100){$results | Out-File -FilePath ".\sharedmailboxpermissions_report.txt" -Encoding UTF8; $results = $null; $count = $null;
 }} else {
            $output = "$($mailbox.DisplayName): $currentUser - No permissions"
        }
        # Output to the console
        Write-Output $output
		  $results += $output
        $results | Out-File -FilePath ".\permissions_report.txt" -Encoding UTF8
        # Store the result for later use
          }if ($count = 100){$results | Out-File -FilePath ".\sharedmailboxpermissions_report.txt" -Encoding UTF8; $results = $null; $count = $null;
 }
 }


# Step 6: Save results to a file
$results | Out-File -FilePath "sharedmailboxpermissions_report.txt" -Encoding UTF8


