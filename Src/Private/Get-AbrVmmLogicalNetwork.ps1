function Get-AbrVmmLogicalNetwork {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Logical Network information
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
                if ($LogicalNetworks = Get-SCLogicalNetworkDefinition | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Networking information."
                    Section -Style Heading2 'Logical Networks' {
                        $VmmLogicalNetworksInfo = @()
                        foreach ($LogicalNetwork in $LogicalNetworks) {
                            $InObj = [Ordered]@{
                                'Name' = $LogicalNetwork.LogicalNetwork
                                'ID' = $LogicalNetwork.ID
                                'Isolation Type' = $LogicalNetwork.IsolationType
                                'Host Groups' = $LogicalNetwork.HostGroups -join ', '
                            }

                            $VmmLogicalNetworksInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.Networking -ge 2) {
                            Paragraph "The following sections detail the configuration of the logical networks."
                            foreach ($LogicalNetwork in $VmmLogicalNetworksInfo) {
                                Section -Style Heading3 "$($LogicalNetwork.Name)" {
                                    $TableParams = @{
                                        Name = "Logical Networks - $($LogicalNetwork.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $LogicalNetwork | Table @TableParams
                                    Get-AbrVmmVlanSubnet -SubnetVLans ($LogicalNetworks | Where-Object { $_.ID -eq $LogicalNetwork.ID }).SubnetVLans
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the logical networks."
                            BlankLine
                            $TableParams = @{
                                Name = "Logical Networks - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'Isolation Type', 'Host Groups'
                                ColumnWidths = 40, 30, 30
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmLogicalNetworksInfo | Table @TableParams
                            Get-AbrVmmVlanSubnet -SubnetVLans $LogicalNetworks.SubnetVLans

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