. $PSScriptRoot\Test-PermissionJob.ps1

function Test-Permission {
    [CmdletBinding()]
    param (
        [Parameter()]
        [PSCustomObject]
        $FolderData
    )

    begin {
        function New-TestPermissionResult {
            [CmdletBinding()]
            param (
                [Parameter(Position = 0)]
                [object]
                $BadPermission,

                [Parameter(Position = 1)]
                [string]
                $ActionRequired
            )

            process {
                [PSCustomObject]@{
                    TestName       = "Permission"
                    ResultType     = "BadPermission"
                    Severity       = "Error"
                    Data           = $BadPermission
                    ActionRequired = $ActionRequired
                }
            }
        }

        $startTime = Get-Date
        $badPermissions = @()
    }

    process {
        $folderData.IpmSubtreeByMailbox | ForEach-Object {
            $argumentList = $FolderData.MailboxToServerMap[$_.Name], $_.Name, $_.Group
            $name = $_.Name
            $scriptBlock = ${Function:Test-BadPermissionJob}
            Add-JobQueueJob @{
                ArgumentList = $argumentList
                Name         = "$name Permissions Check"
                ScriptBlock  = $scriptBlock
            }
        }

        $completedJobs = Wait-QueuedJob
        foreach ($job in $completedJobs) {
            if ($job.BadPermissions.Count -gt 0) {
                foreach ($permission in $job.BadPermissions) {
                    New-TestPermissionResult -BadPermission $permission -ActionRequired "Remove this permission."
                }
            }
        }
    }

    end {
        [PSCustomObject]@{
            TestName       = "Permission"
            ResultType     = "Duration"
            Severity       = "Information"
            Data           = ((Get-Date) - $startTime)
            ActionRequired = $null
        }
    }
}
