function Write-TestMailEnabledFolderResult {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object]
        $TestResult
    )

    begin {
        $results = [System.Collections.ArrayList]::new()
    }

    process {
        $results += $TestResult
    }

    end {
        if ($results.Count -gt 0) {
            $byResultType = $results | Group-Object ResultType
            foreach ($group in $byResultType) {
                if ($group.Name -eq "MailEnabledSystemFolder") {
                    Write-Host
                    Write-Host $group.Count "system folders are mail-enabled. These folders should be mail-disabled."
                } elseif ($group.Name -eq "MailEnabledWithNoADObject") {
                    Write-Host
                    Write-Host $group.Count "folders are mail-enabled, but have no AD object. These folders should be mail-disabled."
                } elseif ($group.Name -eq "MailDisabledWithProxyGuid") {
                    Write-Host
                    Write-Host $group.Count "folders are mail-disabled, but have proxy GUID values. These folders should be mail-enabled."
                } elseif ($group.Name -eq "OrphanedMPF") {
                    Write-Host
                    Write-Host $group.Count "mail public folders are orphaned. These directory objects should be deleted."
                } elseif ($group.Name -eq "OrphanedMPFDuplicate") {
                    Write-Host
                    Write-Host $group.Count "mail public folders point to public folders that point to a different directory object. These should be deleted. Their email addresses may be merged onto the linked object."
                } elseif ($group.Name -eq "OrphanedMPFDisconnected") {
                    Write-Host
                    Write-Host $group.Count "mail public folders point to public folders that are mail-disabled. These require manual intervention. Either the directory object should be deleted, or the folder should be mail-enabled, or both."
                }
            }
        }
    }
}
