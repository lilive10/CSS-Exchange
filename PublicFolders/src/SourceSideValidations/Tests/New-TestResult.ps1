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
        [ValidateSet("Information", "Warning", "Error")]
        [string]
        $Severity,

        [Parameter(Mandatory = $true)]
        [string]
        $FolderIdentity,

        [Parameter(Mandatory = $true)]
        [string]
        $FolderEntryId,

        [Parameter(Mandatory = $false)]
        [string]
        $ResultData
    )

    [PSCustomObject]@{
        TestName       = $TestName
        ResultType     = $ResultType
        Severity       = $Severity
        FolderIdentity = $FolderIdentity
        FolderEntryId  = $FolderEntryId
        ResultData     = $ResultData
    }
}