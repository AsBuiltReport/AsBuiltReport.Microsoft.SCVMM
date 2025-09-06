function Get-AbrVmmHostGroup {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Host Summary information
    .DESCRIPTION

    .NOTES
        Version:        0.1.1
        Author:         AsBuiltReport Organization
        Twitter:        @AsBuiltReport
        Github:         AsBuiltReport
    .EXAMPLE

    .LINK

    #>
    [CmdletBinding()]
    param (
    )

    begin {
        Write-PScriboMessage "Hosts InfoLevel set at $($InfoLevel.Hosts)."
    }

    process {
        try {
            if ($InfoLevel.HostGroups -gt 0) {
                if ($VMHostGroups = Get-SCVMHostGroup -VMMServer $ConnectVmmServer | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Host information."
                    Section -Style Heading2 'Host Groups' {
                        $VmmHostGroupsInfo = @()
                        foreach ($VMHostGroup in $VMHostGroups) {
                            $InObj = [Ordered]@{
                                'Name' = $VMHostGroup.Name
                                'Path' = $VMHostGroup.Path
                                'Parent Host Group' = $VMHostGroup.ParentHostGroup
                                'IsRoot' = $VMHostGroup.IsRoot
                                'Inherit Network Settings' = $VMHostGroup.InheritNetworkSettings
                                'Description' = $VMHostGroup.Description
                            }
                            $VmmHostGroupsInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.HostGroups -ge 2) {
                            Paragraph "The following sections detail the configuration of the vmm host groups."
                            foreach ($VMHostGroup in $VmmHostGroupsInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($VMHostGroup.Name)" {
                                    $TableParams = @{
                                        Name = "Host Groups - $($VMHostGroup.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMHostGroup | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the vmm host groups."
                            BlankLine
                            $TableParams = @{
                                Name = "Host Groups - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'Path', 'Parent Host Group'
                                ColumnWidths = 33, 33, 34
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmHostGroupsInfo | Table @TableParams
                        }
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}