Function Import_Csv {
    #Importing UserPrincipalName From The Csv
    Try {
        $UserDetails = @()
        Write-Host "Importing UserPrincipalNames from Csv..."
        $UPNs = Import-Csv $InputCsvFilePath
        foreach ($UPN in $UPNs) {
            $$ = $UPN.User_Principal_Name
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

