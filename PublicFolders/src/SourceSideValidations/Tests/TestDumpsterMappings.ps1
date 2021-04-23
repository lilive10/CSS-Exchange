function Test-DumpsterMapping {
    [CmdletBinding()]
    [OutputType([System.Object[]])]
    param (
        [Parameter()]
        [PSCustomObject]
        $FolderData
    )

    begin {
        $startTime = Get-Date
        $progressCount = 0
        $sw = New-Object System.Diagnostics.Stopwatch
        $sw.Start()
        $progressParams = @{
            Activity = "Checking dumpster mappings"
            Id       = 2
            ParentId = 1
        }
    }

    process {
        $FolderData.IpmSubtree | ForEach-Object {
            $progressCount++
            if ($sw.ElapsedMilliseconds -gt 1000) {
                $sw.Restart()
                Write-Progress @progressParams -Status $progressCount -PercentComplete ($progressCount * 100 / $FolderData.IpmSubtree.Count)
            }

            if (-not (Test-DumpsterValid $_ $FolderData)) {
                New-TestDumpsterMappingResult $_ "The -ExcludeDumpsters switch can be used to skip these folders during migration, or the folders can be deleted."
            }
        }

        Write-Progress @progressParams -Status "Checking EFORMS dumpster mappings"

        $FolderData.NonIpmSubtree | Where-Object { $_.Identity -like "\NON_IPM_SUBTREE\EFORMS REGISTRY\*" } | ForEach-Object {
            if (-not (Test-DumpsterValid $_ $FolderData)) {
                New-TestDumpsterMappingResult $_
            }
        }
    }

    end {
        Write-Progress @progressParams -Completed
        Write-Host "Get-BadDumpsterMappings duration" ((Get-Date) - $startTime)
    }
}

function Test-DumpsterValid {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter()]
        [PSCustomObject]
        $Folder,

        [Parameter()]
        [PSCustomObject]
        $FolderData
    )

    begin {
        $valid = $true
    }

    process {
        $dumpster = $FolderData.NonIpmEntryIdDictionary[$Folder.DumpsterEntryId]

        if ($null -eq $dumpster -or
            (-not $dumpster.Identity.StartsWith("\NON_IPM_SUBTREE\DUMPSTER_ROOT", "OrdinalIgnoreCase")) -or
            $dumpster.DumpsterEntryId -ne $Folder.EntryId) {

            $valid = $false
        }
    }

    end {
        return $valid
    }
}

function New-TestDumpsterMappingResult {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [object]
        $Folder,

        [Parameter(Position = 1)]
        [string]
        $ActionRequired
    )

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
}

function Write-TestDumpsterMappingResult {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object]
        $TestResult
    )

    begin {
        $badDumpsters = [System.Collections.ArrayList]::new()
    }

    process {
        $badDumpsters += $TestResult
    }

    end {
        if ($badDumpsters.Count -gt 0) {
            Write-Host
            Write-Host $badDumpsters.Count "folders have invalid dumpster mappings."
            Write-Host $badDumpsters[0].ActionRequired
        }
    }
}
