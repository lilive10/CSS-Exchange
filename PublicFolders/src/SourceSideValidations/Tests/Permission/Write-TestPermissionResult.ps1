function Write-TestBadPermissionResult {
    [CmdletBinding()]
    param (

    )

    begin {

    }

    process {
        if ($badPermissions.Count -gt 0) {
            $badPermissionsFile = Join-Path $PSScriptRoot "InvalidPermissions.csv"
            $badPermissions | Export-Csv -Path $badPermissionsFile -NoTypeInformation

            Write-Host
            Write-Host $badPermissions.Count "invalid permissions were found. These are listed in the following CSV file:"
            Write-Host $badPermissionsFile -ForegroundColor Green
            Write-Host "The invalid permissions can be removed using the RemoveInvalidPermissions switch as follows:"
            Write-Host ".\SourceSideValidations.ps1 -RemoveInvalidPermissions" -ForegroundColor Green
        }
    }

    end {

    }
}