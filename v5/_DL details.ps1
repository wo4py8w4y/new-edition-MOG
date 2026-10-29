# Import distribution groups from CSV and output details

# Import the CSV file
$groups = Import-Csv -Path ".\DLs.csv"

$total = $groups.Count
$counter = 0

foreach ($group in $groups) {
    $counter++
  

    $dg = Get-DistributionGroup -Identity $group.DisplayName
    $owner = (Get-DistributionGroup -Identity $dg.Identity | Get-DistributionGroupOwner).Name -join '; '
    $members = (Get-DistributionGroupMember -Identity $dg.Identity -resultsize unlimited | Where-Object { $_.RecipientType -eq 'UserMailbox' } | Select-Object -ExpandProperty Alias) -join '; '

    $output = [PSCustomObject]@{
        GroupName    = $dg.DisplayName
        EmailAddress = $dg.PrimarySmtpAddress
        Owner        = $owner
        GUID         = $dg.Guid
        Members      = $members
    }
$MemberCount  = (Get-DistributionGroupMember -Identity $dg.Identity -resultsize unlimited | Where-Object { $_.RecipientType -eq 'UserMailbox' }).Count
$MemberList   = (Get-DistributionGroupMember -Identity $dg.Identity -resultsize unlimited | Where-Object { $_.RecipientType -eq 'UserMailbox' } | Select-Object -ExpandProperty DisplayName) -join '; '
$LastUpdated  = (Get-DistributionGroup -Identity $dg.Identity).WhenChanged
$CreationDate = (Get-DistributionGroup -Identity $dg.Identity).WhenCreated   
$LastAccessed  = (Get-DistributionGroup -Identity $dg.Identity).LastAccessed 


   $xtra =  $MemberCount , $MemberList , $LastUpdated , $CreationDate , $LastAccessed
    $output | Add-Member -MemberType NoteProperty -Name "MemberCount" -Value $xtra[0]
    $output | Add-Member -MemberType NoteProperty -Name "MemberList" -Value $xtra[1]
    $output | Add-Member -MemberType NoteProperty -Name "LastUpdated" -Value $xtra[2]
    $output | Add-Member -MemberType NoteProperty -Name "CreationDate" -Value $xtra[3]
    $output | Add-Member -MemberType NoteProperty -Name "LastAccessed" -Value $xtra[4]
    Write-Progress -Activity "Processing groups" -Status "$counter of $total" -PercentComplete (($counter / $total) * 100)
    # Check if the CSV file has a column named 'DisplayName'    
    $fileName = "C:\Path\To\Your\Output\DL_$($dg.DisplayName)_details.csv"
    $output | Export-Csv -Path $fileName -NoTypeInformation
}