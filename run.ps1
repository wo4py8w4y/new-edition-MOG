Param
(
    [string]$UserName = $Null,
    [string]$Password = $Null,
    [string]$UserPrincipalName = $Null,
    [string]$InputCsvFilePath = $Null
)

Function Import_Csv {
    #Importing UserPrincipalName From The Csv
    Try {
        $UserDetails = @()
        Write-Host "Importing UserPrincipalNames from Csv..."
        $UPNs = Import-Csv $InputCsvFilePath
        foreach ($UPN in $UPNs) {
            $UserPrincipalName = $UPN.User_Principal_Name
            Try {
                Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop | ForEach-Object {
                    List_DLs_That_User_Is_A_Member

                }
            }
            Catch {
                Write-Host "$UserPrincipalName is not a valid user"
            }
        }

    }
    catch {
        Write-Host "$InputCsvFilePath is not a valid file path"
    }
}
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
Function OpenOutputCsv {
    #Open Output File After Execution
    If ((Test-Path $OutputCsv) -eq "True") {
        Write-Host `n"The output file contains:" -NoNewline -ForegroundColor Yellow; Write-Host $ProcessedUserCount users `n
        Write-Host " The Output file available in:" -NoNewline -ForegroundColor Yellow; $OutputCsv
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open output file?", 0, "open output file", 4)
        If ($UserInput -eq 6) {
            Invoke-Item "$OutputCsv"
        }
    }
}
#Connect_Exo
$Global:ProcessedUserCount = 1
$OutputCsv = ".\ListDLs_UsersIsMemberOf_$((Get-Date -format MMM-dd` hh-mm` tt).ToString()).csv"
If ($UserPrincipalName -ne "") {
    Try {
        write-Host "Checking $UserPrincipalName is a valid user or not"
        Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop | ForEach-Object {
            List_DLs_That_User_Is_A_Member
        }
    }
    Catch {
        Write-Host "$UserPrincipalName is not a valid user"
    }
}
Elseif ($InputCsvFilePath -ne "") {
    Import_Csv
}
Else {
    Get-Mailbox -ResultSize unlimited -RecipientTypeDetails UserMailbox | ForEach-Object {
        List_DLs_That_User_Is_A_Member
    }
}
OpenOutputCsv