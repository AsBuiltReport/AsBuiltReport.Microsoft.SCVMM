function Get-AbrVmmVmSubnet {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM VM Subnets information
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
        [Parameter(Mandatory)]
        $VMSubnet
    )

    begin {
        Write-PScriboMessage "Networking InfoLevel set at $($InfoLevel.Networking)."
    }

    process {
        try {
            if ($InfoLevel.Networking -gt 0) {
                if ($VMSubnet) {
                    Write-PScriboMessage "Collecting VMM VM Subnets information."
                    Section -Style Heading5 'VM Subnets' {
                        $VmmVlansSubnetsInfo = @()
                        foreach ($VlansSubnets in ($VMSubnet | Sort-Object -Property Name)) {
                            $InObj = [Ordered]@{
                                'Name' = $VlansSubnets.Name
                                'ID' = $VlansSubnets.ID
                                'VM Network' = $VlansSubnets.VMNetwork
                                'Encrypted' = $VlansSubnets.EncryptionEnabled
                                'Allows Intra Port Communication' = $VlansSubnets.AllowsIntraPortCommunication
                            }

                            $VmmVlansSubnetsInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.Networking -ge 2) {
                            Paragraph "The following sections detail the configuration of the vm subnets."
                            foreach ($VlansSubnets in $VmmVlansSubnetsInfo) {
                                Section -Style Heading6 "$($VlansSubnets.Name)" {
                                    $TableParams = @{
                                        Name = "VM Subnets - $($VlansSubnets.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VlansSubnets | Table @TableParams
                                    Get-AbrVmmVmVlanSubnet -Name $VlansSubnets.Name -SubnetVLans ($VMSubnet | Where-Object { $_.ID -eq $VlansSubnets.ID }).SubnetVLans
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the vm subnets."
                            BlankLine
                            $TableParams = @{
                                Name = "VM Subnets - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'VM Network', 'Encrypted', 'Allows Intra Port Communication'
                                ColumnWidths = 25, 25, 25, 25
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmVlansSubnetsInfo | Table @TableParams
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