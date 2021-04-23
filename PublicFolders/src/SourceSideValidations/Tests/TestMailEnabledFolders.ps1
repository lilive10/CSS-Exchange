function Test-MailEnabledFolder {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
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
            Activity = "Validating mail-enabled public folders"
            Id       = 2
            ParentId = 1
        }
    }

    process {
        $FolderData.NonIpmSubtree | Where-Object { $_.MailEnabled -eq $true } | ForEach-Object { New-TestMailEnabledFolderResult $_.Identity $_.EntryId "MailEnabledSystemFolder" "Run Disable-MailPublicFolder on this folder" }
        $ipmSubtreeMailEnabled = @($FolderData.IpmSubtree | Where-Object { $_.MailEnabled -eq $true })
        $mailDisabledWithProxyGuid = @($FolderData.IpmSubtree | Where-Object { $_.MailEnabled -ne $true -and -not [string]::IsNullOrEmpty($_.MailRecipientGuid) -and [Guid]::Empty -ne $_.MailRecipientGuid } | ForEach-Object { $_.Identity.ToString() })
        $mailDisabledWithProxyGuid | ForEach-Object {
            $params = @{
                Identity = $_.Identity
                EntryId = $_.EntryId
                ResultType = "MailDisabledWithProxyGuid"
                ActionRequired = "Run Enable-MailPublicFolder on this folder. It can be mail-disabled again afterwards if desired."
            }

            New-TestMailEnabledFolderResult @params
        }


        $mailPublicFoldersLinked = New-Object 'System.Collections.Generic.Dictionary[string, object]'
        $progressParams.CurrentOperation = "Checking for missing AD objects"
        $startTimeForThisCheck = Get-Date
        for ($i = 0; $i -lt $ipmSubtreeMailEnabled.Count; $i++) {
            $progressCount++
            if ($sw.ElapsedMilliseconds -gt 1000) {
                $sw.Restart()
                $elapsed = ((Get-Date) - $startTimeForThisCheck)
                $estimatedRemaining = [TimeSpan]::FromTicks($ipmSubtreeMailEnabled.Count / $progressCount * $elapsed.Ticks - $elapsed.Ticks).ToString("hh\:mm\:ss")
                Write-Progress @progressParams -PercentComplete ($i * 100 / $ipmSubtreeMailEnabled.Count) -Status ("$i of $($ipmSubtreeMailEnabled.Count) Estimated time remaining: $estimatedRemaining")
            }
            $result = Get-MailPublicFolder $ipmSubtreeMailEnabled[$i].Identity -ErrorAction SilentlyContinue
            if ($null -eq $result) {
                $params = @{
                    Identity = $ipmSubtreeMailEnabled[$i].Identity
                    EntryId = $ipmSubtreeMailEnabled[$i].EntryId
                    ResultType = "MailEnabledWithNoADObject"
                    ActionRequired = "Run Disable-MailPublicFolder on this folder"
                }

                New-TestMailEnabledFolderResult @params
            } else {
                $guidString = $result.Guid.ToString()
                if (-not $mailPublicFoldersLinked.ContainsKey($guidString)) {
                    $mailPublicFoldersLinked.Add($guidString, $result) | Out-Null
                }
            }
        }

        $progressCount = 0
        $progressParams.CurrentOperation = "Getting all MailPublicFolder objects"
        $allMailPublicFolders = @(Get-MailPublicFolder -ResultSize Unlimited | ForEach-Object {
                $progressCount++
                if ($sw.ElapsedMilliseconds -gt 1000) {
                    $sw.Restart()
                    Write-Progress @progressParams -Status "$progressCount"
                }

                $_
            })


        $progressCount = 0
        $progressParams.CurrentOperation = "Checking for orphaned MailPublicFolders"
        $orphanedMailPublicFolders = @($allMailPublicFolders | ForEach-Object {
                $progressCount++
                if ($sw.ElapsedMilliseconds -gt 1000) {
                    $sw.Restart()
                    Write-Progress @progressParams -PercentComplete ($progressCount * 100 / $allMailPublicFolders.Count) -Status ("$progressCount of $($allMailPublicFolders.Count)")
                }

                if (!($mailPublicFoldersLinked.ContainsKey($_.Guid.ToString()))) {
                    $_
                }
            })


        $progressParams.CurrentOperation = "Building EntryId HashSets"
        Write-Progress @progressParams
        $byEntryId = New-Object 'System.Collections.Generic.Dictionary[string, object]'
        $FolderData.IpmSubtree | ForEach-Object { $byEntryId.Add($_.EntryId.ToString(), $_) }
        $byPartialEntryId = New-Object 'System.Collections.Generic.Dictionary[string, object]'
        $FolderData.IpmSubtree | ForEach-Object { $byPartialEntryId.Add($_.EntryId.ToString().Substring(44), $_) }

        $progressParams.CurrentOperation = "Checking for orphans that point to a valid folder"
        for ($i = 0; $i -lt $orphanedMailPublicFolders.Count; $i++) {
            if ($sw.ElapsedMilliseconds -gt 1000) {
                $sw.Restart()
                Write-Progress @progressParams -PercentComplete ($i * 100 / $orphanedMailPublicFolders.Count) -Status ("$i of $($orphanedMailPublicFolders.Count)")
            }

            $thisMPF = $orphanedMailPublicFolders[$i]
            $pf = $null
            if ($null -ne $thisMPF.ExternalEmailAddress -and $thisMPF.ExternalEmailAddress.ToString().StartsWith("expf")) {
                $partialEntryId = $thisMPF.ExternalEmailAddress.ToString().Substring(5).Replace("-", "")
                $partialEntryId += "0000"
                if ($byPartialEntryId.TryGetValue($partialEntryId, [ref]$pf)) {
                    if ($pf.MailEnabled -eq $true) {

                        $command = GetCommandToMergeEmailAddresses $pf $thisMPF

                        $params = @{
                            Identity = $thisMPF.DistinguishedName.Replace("/", "\/")
                            EntryId = ""
                            ResultType = "OrphanedMPFDuplicate"
                            ActionRequired = "Delete this directory object"
                        }

                        if ($null -ne $command) {
                            $params.ActionRequired += ", then run the following command to merge the email addresses onto the remaining object:`n`n$command"
                        }

                        New-TestMailEnabledFolderResult @params
                    } else {
                        $params = @{
                            Identity = $thisMPF.DistinguishedName.Replace("/", "\/")
                            EntryId = ""
                            ResultType = "OrphanedMPFDisconnected"
                            ActionRequired = "This requires manual intervention. Either the directory object should be deleted, or the public folder should be mail-enabled, or both."
                        }

                        New-TestMailEnabledFolderResult @params
                    }

                    continue
                }
            }

            if ($null -ne $thisMPF.EntryId -and $byEntryId.TryGetValue($thisMPF.EntryId.ToString(), [ref]$pf)) {
                if ($pf.MailEnabled -eq $true) {

                    $command = GetCommandToMergeEmailAddresses $pf $thisMPF

                    $params = @{
                        Identity = $thisMPF.DistinguishedName.Replace("/", "\/")
                        EntryId = ""
                        ResultType = "OrphanedMPFDuplicate"
                        ActionRequired = "Delete this directory object"
                    }

                    if ($null -ne $command) {
                        $params.ActionRequired += ", then run the following command to merge the email addresses onto the remaining object:`n`n$command"
                    }

                    New-TestMailEnabledFolderResult @params
                } else {
                    $params = @{
                        Identity = $thisMPF.DistinguishedName.Replace("/", "\/")
                        EntryId = ""
                        ResultType = "OrphanedMPFDisconnected"
                        ActionRequired = "This requires manual intervention. Either the directory object should be deleted, or the public folder should be mail-enabled, or both."
                    }

                    New-TestMailEnabledFolderResult @params
                }
            } else {
                $params = @{
                    Identity = $thisMPF.DistinguishedName.Replace("/", "\/")
                    EntryId = ""
                    ResultType = "OrphanedMPF"
                    ActionRequired = "Delete this directory object"
                }

                New-TestMailEnabledFolderResult @params
            }
        }
    }

    end {
        Write-Progress @progressParams -Completed
    }
}

function GetCommandToMergeEmailAddresses($publicFolder, $orphanedMailPublicFolder) {
    $linkedMailPublicFolder = Get-PublicFolder $publicFolder.Identity | Get-MailPublicFolder
    $emailAddressesOnGoodObject = @($linkedMailPublicFolder.EmailAddresses | Where-Object { $_.ToString().StartsWith("smtp:", "OrdinalIgnoreCase") } | ForEach-Object { $_.ToString().Substring($_.ToString().IndexOf(':') + 1) })
    $emailAddressesOnBadObject = @($orphanedMailPublicFolder.EmailAddresses | Where-Object { $_.ToString().StartsWith("smtp:", "OrdinalIgnoreCase") } | ForEach-Object { $_.ToString().Substring($_.ToString().IndexOf(':') + 1) })
    $emailAddressesToAdd = $emailAddressesOnBadObject | Where-Object { -not $emailAddressesOnGoodObject.Contains($_) }
    $emailAddressesToAdd = $emailAddressesToAdd | ForEach-Object { "`"" + $_ + "`"" }
    if ($emailAddressesToAdd.Count -gt 0) {
        $emailAddressesToAddString = [string]::Join(",", $emailAddressesToAdd)
        $command = "Get-PublicFolder `"$($publicFolder.Identity)`" | Get-MailPublicFolder | Set-MailPublicFolder -EmailAddresses @{add=$emailAddressesToAddString}"
        return $command
    } else {
        return $null
    }
}

function New-TestMailEnabledFolderResult {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [string]
        $Identity,

        [Parameter(Position = 1)]
        [string]
        $EntryId,

        [Parameter(Position = 2)]
        [ValidateSet("MailEnabledSystemFolder", "MailEnabledWithNoADObject", "MailDisabledWithProxyGuid", "OrphanedMPF", "OrphanedMPFDuplicate", "OrphanedMPFDisconnected")]
        [string]
        $ResultType,

        [Parameter(Position = 3)]
        [string]
        $ActionRequired
    )

    process {
        [PSCustomObject]@{
            TestName        = "MailEnabledFolder"
            ResultType      = $ResultType
            Identity        = $Identity
            EntryId         = $EntryId
            $ActionRequired = $ActionRequired
        }
    }
}

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

function Update-MailEnabledFolderResult {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object]
        $TestResult
    )

    process {
        if ($TestResult.ResultType -eq "MailEnabledSystemFolder") {
            Write-Host
            Write-Host $group.Count "system folders are mail-enabled. These folders should be mail-disabled."
        } elseif ($TestResult.ResultType -eq "MailEnabledWithNoADObject") {
            Write-Host
            Write-Host $group.Count "folders are mail-enabled, but have no AD object. These folders should be mail-disabled."
        } elseif ($TestResult.ResultType -eq "MailDisabledWithProxyGuid") {
            Write-Host
            Write-Host $group.Count "folders are mail-disabled, but have proxy GUID values. These folders should be mail-enabled."
        } elseif ($TestResult.ResultType -eq "OrphanedMPF") {
            Write-Host
            Write-Host $group.Count "mail public folders are orphaned. These directory objects should be deleted."
        } elseif ($TestResult.ResultType -eq "OrphanedMPFDuplicate") {
            Write-Host
            Write-Host $group.Count "mail public folders point to public folders that point to a different directory object. These should be deleted. Their email addresses may be merged onto the linked object."
        } elseif ($TestResult.ResultType -eq "OrphanedMPFDisconnected") {
            Write-Host
            Write-Host $group.Count "mail public folders point to public folders that are mail-disabled. These require manual intervention. Either the directory object should be deleted, or the folder should be mail-enabled, or both."
        }
    }
}
