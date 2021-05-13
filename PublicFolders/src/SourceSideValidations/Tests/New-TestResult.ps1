function New-TestResult {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $TestName,

        [Parameter(Mandatory = $true)]
        [string]
        $ResultType,

        [Parameter(Mandatory = $true)]
        [string]
        $Severity,

        [Parameter(Mandatory = $true)]
        [string]
        $FolderIdentity,

        [Parameter(Mandatory = $true)]
        [string]
        $FolderEntryId,
    )

    begin {

    }

    process {
        [PSCustomObject]@{
            TestName       = "DumpsterMapping"
            ResultType     = "BadDumpsterMapping"
            Severity       = "Error"
            Data           = [PSCustomObject]@{
                Identity = $Folder.Identity
                EntryId  = $Folder.EntryId
            }
            ActionRequired = $ActionRequired
        }
    }

    end {

    }
}