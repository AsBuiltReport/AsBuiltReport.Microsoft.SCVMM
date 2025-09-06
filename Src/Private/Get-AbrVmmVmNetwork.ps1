function Get-AbrVmmVmNetwork {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM VM Network information
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
                if ($VMNetworks = Get-SCVMNetwork | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM VM Networks information."
                    Section -Style Heading3 'VM Networks' {
                        $VmmVMNetworksInfo = @()
                        foreach ($VMNetwork in $VMNetworks) {
                            $InObj = [Ordered]@{
                                'Name' = $VMNetwork.Name
                                'ID' = $VMNetwork.ID
                                'Logical Network' = $VMNetwork.LogicalNetwork
                                'VM Subnets' = $VMNetwork.VMSubnet.Name -join ', '
                                'Isolation Type' = $VMNetwork.IsolationType
                                'Enabled' = $VMNetwork.Enabled
                                'Is Assigned?' = $VMNetwork.IsAssigned
                                'Description' = $VMNetwork.Description
                            }

                            $VmmVMNetworksInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.Networking -ge 2) {
                            Paragraph "The following sections detail the configuration of the vm networks."
                            foreach ($VMNetwork in $VmmVMNetworksInfo) {
                                Section -Style Heading4 "$($VMNetwork.Name)" {
                                    $TableParams = @{
                                        Name = "VM Networks - $($VMNetwork.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMNetwork | Table @TableParams
                                    Get-AbrVmmVmSubnet -VMSubnet ($VMNetworks | Where-Object { $_.ID -eq $VMNetwork.ID }).VMSubnet
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the vm networks."
                            BlankLine
                            $TableParams = @{
                                Name = "VM Networks - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'Logical Network', 'VM Subnets', 'Isolation Type', 'Enabled'
                                ColumnWidths = 20, 20, 20, 25, 15
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmVMNetworksInfo | Table @TableParams
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