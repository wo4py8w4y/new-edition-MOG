$groups1 = Import-Csv -Path ".\DLgroups.csv"
$groups = $groups1.Identity
$users1 = Import-Csv -Path ".\Usermailboxes-test.csv"
$users = $users1.primarysmtpaddress

# Step 3: Prepare output file
$outputFile = ".\DLmembershipResults.txt"
$count = 0
#Clear-Content -Path $outputFile -ErrorAction SilentlyContinue # Clear previous results

# Step 4: Check membership for each user in each distribution group
foreach ($user in $users) {
    $GUID = Get-Recipient $user
	$count + 1
    foreach ($group in $groups) {
        # Retrieve group members
                $isMember = Get-DistributionGroupMember -Identity $group -resultsize 50| Where-Object {
            $_.PrimarySmtpAddress -eq $GUID.PrimarySmtpAddress -or $_.GUID -eq $GUID.GUID
        }

        # Check if user is a member
        if ($isMember) {
            $resultMessage = "$user is a member of $group"
            Write-Output $resultMessage
            $resultMessage | Out-File -FilePath ".\DLmembershipresults.txt" -Append
        } else {
            $resultMessage = "$user is NOT a member of $group"
            Write-Output $resultMessage
            $resultMessage | Out-File -FilePath ".\DLmembershipresults.txt" -Append
        }if ($count = '100'){out-file -FilePath ".\DLmembershipresultsdump.txt" -Append; $count = 0; $resultMessage = $null}
    }
}

# Disconnect from Exchange Online
#Disconnect-ExchangeOnline -Confirm:$false

