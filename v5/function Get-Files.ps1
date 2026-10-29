function Get-Files
{
    [CmdletBinding()]
    Param
    (
        [validateset("Name","Length")]
        [string]$Sort = "Length",

        [switch]$Descending
    )

    $splat = @{}

    if ($Sort -eq "Length")
    {
        $splat.Add( 'Descending', $true )
    }

    if ( $Descending -eq $true ) 
    { 
        if ( !$splat.ContainsKey("Descending") )
        {
            $splat.Add( 'Descending', $true )
        }
    }

    Get-ChildItem -Recurse -File | 
        Sort-Object $($sort) @splat |
            Format-Table Name, @{ n="Size (GB)"; e={$_.Length/1GB}; a="right"; f="{0:N2}" }
}

