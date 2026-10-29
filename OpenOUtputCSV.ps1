Function OpenOutputCsv {
    #Open Output File After Execution
    If ((Test-Path $OutputCsv) -eq "True") {
        Write-Host `n"The output file contains:" -NoNewline -ForegroundColor Yellow; Write-Host $ProcessedUserCount users `n
        Write-Host " The Output file available in:" -NoNewline -ForegroundColor Yellow; $OutputCsv
        Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
        Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `n`n
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open output file?", 0, "open output file", 4)
        If ($UserInput -eq 6) {
            Invoke-Item "$OutputCsv"
        }
    }
}

