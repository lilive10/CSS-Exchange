[CmdletBinding(DefaultParameterSetName = "Default", SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $false, ParameterSetName = "Default")]
    [bool]
    $StartFresh = $true,

    [Parameter(Mandatory = $true, ParameterSetName = "RemoveInvalidPermissions")]
    [Switch]
    $RemoveInvalidPermissions,

    [Parameter(ParameterSetName = "RemoveInvalidPermissions")]
    [string]
    $CsvFile = (Join-Path $PSScriptRoot "InvalidPermissions.csv"),

    [Parameter()]
    [switch]
    $SkipVersionCheck
)

. $PSScriptRoot\Get-FolderData.ps1
. $PSScriptRoot\Get-LimitsExceeded.ps1
. $PSScriptRoot\JobQueue.ps1
. $PSScriptRoot\Remove-InvalidPermission.ps1
. $PSScriptRoot\..\..\..\Shared\Test-ScriptVersion.ps1

if (-not $SkipVersionCheck) {
    if (Test-ScriptVersion -AutoUpdate) {
        # Update was downloaded, so stop here.
        Write-Host "Script was updated. Please rerun the command."
        return
    }
}

if ($RemoveInvalidPermissions) {
    if (-not (Test-Path $CsvFile)) {
        Write-Error "File not found: $CsvFile"
    } else {
        Remove-InvalidPermission -CsvFile $CsvFile
    }
    return
}

$startTime = Get-Date

$startingErrorCount = $Error.Count

Set-ADServerSettings -ViewEntireForest $true

if ($Error.Count -gt $startingErrorCount) {
    # If we already have errors, we're not running from the right shell.
    return
}

$progressParams = @{
    Activity = "Validating public folders"
    Id       = 1
}

Write-Progress @progressParams -Status "Step 1 of 5"

$folderData = Get-FolderData -StartFresh $StartFresh

if ($folderData.IpmSubtree.Count -lt 1) {
    return
}

$script:anyDatabaseDown = $false
Get-Mailbox -PublicFolder | ForEach-Object {
    try {
        $db = Get-MailboxDatabase $_.Database -Status
        if ($db.Mounted) {
            $folderData.MailboxToServerMap[$_.DisplayName] = $db.Server
        } else {
            Write-Error "Database $db is not mounted. This database holds PF mailbox $_ and must be mounted."
            $script:anyDatabaseDown = $true
        }
    } catch {
        Write-Error $_
        $script:anyDatabaseDown = $true
    }
}

if ($script:anyDatabaseDown) {
    Write-Host "One or more PF mailboxes cannot be reached. Unable to proceed."
    return
}

# Now we're ready to do the checks

Write-Progress @progressParams -Status "Step 2 of 5"

$badDumpsters = @(Test-DumpsterMapping -FolderData $folderData)

Write-Progress @progressParams -Status "Step 3 of 5"

$limitsExceeded = Get-LimitsExceeded -FolderData $folderData

Write-Progress @progressParams -Status "Step 4 of 5"

$badMailEnabled = Get-BadMailEnabledFolder -FolderData $folderData

Write-Progress @progressParams -Status "Step 5 of 5"

$badPermissions = @(Test-BadPermission -FolderData $folderData)

# Output the results

$badMailEnabled | Write-TestMailEnabledFolderResult

$badDumpsters | Write-TestDumpsterMappingResult

if ($limitsExceeded.ChildCount.Count -gt 0) {
    $tooManyChildFoldersFile = Join-Path $PSScriptRoot "TooManyChildFolders.txt"
    Set-Content -Path $tooManyChildFoldersFile -Value $limitsExceeded.ChildCount

    Write-Host
    Write-Host $limitsExceeded.ChildCount.Count "folders have exceeded the child folder limit of 10,000. These folders are"
    Write-Host "listed in the following file:"
    Write-Host $tooManyChildFoldersFile -ForegroundColor Green
    Write-Host "Under each of the listed folders, child folders should be relocated or deleted to reduce this number."
}

if ($limitsExceeded.FolderPathDepth.Count -gt 0) {
    $pathTooDeepFile = Join-Path $PSScriptRoot "PathTooDeep.txt"
    Set-Content -Path $pathTooDeepFile -Value $limitsExceeded.FolderPathDepth

    Write-Host
    Write-Host $limitsExceeded.FolderPathDepth.Count "folders have exceeded the path depth limit of 299. These folders are"
    Write-Host "listed in the following file:"
    Write-Host $pathTooDeepFile -ForegroundColor Green
    Write-Host "These folders should be relocated to reduce the path depth, or deleted."
}

if ($limitsExceeded.ItemCount.Count -gt 0) {
    $tooManyItemsFile = Join-Path $PSScriptRoot "TooManyItems.txt"
    Set-Content -Path $tooManyItemsFile -Value $limitsExceeded.ItemCount

    Write-Host
    Write-Host $limitsExceeded.ItemCount.Count "folders exceed the maximum of 1 million items. These folders are listed"
    Write-Host "in the following file:"
    Write-Host $tooManyItemsFile
    Write-Host "In each of these folders, items should be deleted to reduce the item count."
}

$badPermissions | Write-TestBadPermissionResult



$folderCountMigrationLimit = 250000

if ($folderData.IpmSubtree.Count -gt $folderCountMigrationLimit) {
    Write-Host
    Write-Host "There are $($folderData.IpmSubtree.Count) public folders in the hierarchy. This exceeds"
    Write-Host "the supported migration limit of $folderCountMigrationLimit for Exchange Online. The number"
    Write-Host "of public folders must be reduced prior to migrating to Exchange Online."
} elseif ($folderData.IpmSubtree.Count * 2 -gt $folderCountMigrationLimit) {
    Write-Host
    Write-Host "There are $($folderData.IpmSubtree.Count) public folders in the hierarchy. Because each of these"
    Write-Host "has a dumpster folder, the total number of folders to migrate will be $($folderData.IpmSubtree.Count * 2)."
    Write-Host "This exceeds the supported migration limit of $folderCountMigrationLimit for Exchange Online."
    Write-Host "New-MigrationBatch can be run with the -ExcludeDumpsters switch to skip the dumpster"
    Write-Host "folders, or public folders may be deleted to reduce the number of folders."
}

$private:endTime = Get-Date

Write-Host
Write-Host "SourceSideValidations complete. Total duration" ($endTime - $startTime)
