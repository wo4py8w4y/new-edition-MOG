$groups = Import-Csv -Path ".\DLs.csv"

# Check if the CSV file has a column named 'DistributionGroup'
if (-not $groups[0].PSObject.Properties.Name -contains 'DistributionGroup') {
    Write-Error "The CSV file must contain a column named 'DistributionGroup'."
    exit
}

$total = $groups.Count
$counter = 0

foreach ($grp in $groups) {
    $counter++
    Write-Progress -Activity "Processing Distribution Groups" -Status "Processing $($grp.DistributionGroup) ($counter of $total)" -PercentComplete (($counter / $total) * 100)

    if (-not $grp.DistributionGroup) {
        Write-Warning "Missing DistributionGroup in row $counter. Skipping."
        continue
    }

    $dg = Get-DistributionGroup -Identity $grp.DistributionGroup -ErrorAction SilentlyContinue
    if (-not $dg) {
        Write-Warning "Distribution group not found: $($grp.DistributionGroup). Skipping."
        continue
    } else {
        Write-Host "Processing Distribution Group: $($dg.DistributionGroup)"
    }

    # Get owner (ManagedBy property, if any)
    $owner = $null
    $ownerName = ""
    try {
        if ($dg.ManagedBy) {
            $owner = $dg.ManagedBy | Get-Recipient -ErrorAction Stop
            $ownerName = ($owner | Select-Object -ExpandProperty DisplayName) -join ', '
        }
    } catch {
        Write-Warning "Could not resolve owner for group $($dg.DistributionGroup)."
        $ownerName = ""
    }

    # Get group members
    $members = @()
    try {
        $members = Get-DistributionGroupMember -Identity $dg.Identity -ErrorAction Stop -resultsize unlimited | 
            Select-Object -ExpandProperty PrimarySmtpAddress
    } catch {
        Write-Warning "Could not retrieve members for $($dg.DistributionGroup)."
    }

    # Prepare output object
    $output = [PSCustomObject]@{
        GroupDistributionGroup     = $dg.DistributionGroup
        PrimarySMTPAddress         = $dg.PrimarySmtpAddress
        Alias                      = $dg.Alias
        GUID                       = $dg.Guid
        Owner                      = $ownerName
        Members                    = ($members -join ', ')
    }

    # Export to a separate CSV for each group
    $safeName = ($dg.alias -replace '[\\\/:*?"<>|]', '_')
    if ($safeName.Length -gt 255) {
        $safeName = $safeName.Substring(0, 255)  # Limiting to 255 characters
    }
    $fileName = "DL_$safeName.csv"
    try {
        $output | Export-Csv -Path $fileName -NoTypeInformation
        Write-Host "Exported $fileName successfully."
    } catch {
        Write-Warning "Failed to export CSV for $($dg.DistributionGroup)."
    }
}

# Clear the progress bar after completion
Write-Progress -Activity "Processing Distribution Groups" -Completed -Status "Done"