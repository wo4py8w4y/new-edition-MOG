Function List_DLs_That_User_Is_A_Member {
    #Finding Distribution List that  User is a Member
    $Result = @()
    $DistinguishedName = $_.DistinguishedName
    $Filter = "Members -Like ""$DistinguishedName"""
    $UserPrincipalName = $_.UserPrincipalName
    $UserDisplayName = $_.DisplayName
    Write-Progress -Activity "Find Distribution Lists that user is a member" -Status "Processed User Count: $Global:ProcessedUserCount" -CurrentOperation "Currently Processing in  $UserPrincipalName"
    $DLs = Get-DistributionGroup -ResultSize Unlimited -Filter $Filter
    $GroupCount = $DLs | Measure-Object | Select-Object count
    If ($GroupCount.count -ne 0) {
        $DLsCount = $GroupCount.count
        $DLsName = $DLs.Name
        $DLsEmailAddress = $DLs.PrimarySmtpAddress
    }
    Else {
        $DLsName = "-"
        $DlsEmailAddress = "-"
        $DLsCount = '0'
    }
    $Result = New-Object PsObject -Property @{'User Principal Name' = $UserPrincipalName; 'User Display Name' = $UserDisplayName; 'No of DLs that user is a member' = $DLsCount; 'DLs Name' = $DLsName -join ','; 'DLs Email Adddress' = $DLsEmailAddress -join ','; } 
    $Result | Select-Object 'User Principal Name', 'User Display Name', 'No Of DLs That User Is A Member', 'DLs Name', 'DLs Email Adddress' | Export-Csv  $OutputCsv -NoTypeInformatio -Append 
    $Global:ProcessedUserCount++
}

