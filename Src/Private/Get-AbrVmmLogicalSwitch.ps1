function Get-AbrVmmLogicalSwitch {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Logical Switch information
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
        Write-PScriboMessage "Networking InfoLevel set at $($InfoLevel.Networking)."
    }

    process {
        try {
            if ($InfoLevel.Networking -gt 0) {
                if ($LogicalSwitches = Get-SCLogicalSwitch | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Logical Switches information."
                    Section -Style Heading3 'Logical Switches' {
                        $VmmLogicalSwitchesInfo = @()
                        foreach ($LogicalSwitch in $LogicalSwitches) {
                            $InObj = [Ordered]@{
                                'Name' = $LogicalSwitch.Name
                                'ID' = $LogicalSwitch.ID
                                'Uplink Mode' = $LogicalSwitch.UplinkMode
                                'Minimum Bandwidth Mode' = $LogicalSwitch.MinimumBandwidthMode
                                'Enable SR-IOV' = $LogicalSwitch.EnableSriov
                                'Description' = $LogicalSwitch.Description
                            }

                            $VmmLogicalSwitchesInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.Networking -ge 2) {
                            Paragraph "The following sections detail the configuration of the logical switches."
                            foreach ($LogicalSwitch in $VmmLogicalSwitchesInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($LogicalSwitch.Name)" {
                                    $TableParams = @{
                                        Name = "Logical Switches - $($LogicalSwitch.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $LogicalSwitch | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the logical switches."
                            BlankLine
                            $TableParams = @{
                                Name = "Logical Switches - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'Uplink Mode', 'Minimum Bandwidth Mode', 'Enable SR-IOV'
                                ColumnWidths = 25, 25, 25, 25
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmLogicalSwitchesInfo | Table @TableParams
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