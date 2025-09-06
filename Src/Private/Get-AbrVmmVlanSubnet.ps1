function Get-AbrVmmVlanSubnet {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM VLANs and Subnets information
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
        $SubnetVLans
    )

    begin {
        Write-PScriboMessage "Networking InfoLevel set at $($InfoLevel.Networking)."
    }

    process {
        try {
            if ($InfoLevel.Networking -gt 0) {
                if ($SubnetVLans) {
                    Write-PScriboMessage "Collecting VMM VLANs and Subnets information."
                    Section -Style Heading2 'Subnets & VLANs' {
                        $VmmVlansSubnetsInfo = @()
                        foreach ($VlansSubnets in ($SubnetVLans | Sort-Object -Property Name)) {
                            $InObj = [Ordered]@{
                                'Name' = $VlansSubnets.Subnet
                                'VLAN ID' = $VlansSubnets.VLanID
                                'Secondary VLan ID' = $VlansSubnets.SecondaryVLanID -join ', '
                                'Supports DHCP' = $VlansSubnets.SupportsDHCP
                                'Is Assigned To VM Subnet?' = $VlansSubnets.IsAssignedToVMSubnet
                                'Enabled' = $VlansSubnets.IsVLanEnabled
                            }

                            $VmmVlansSubnetsInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.Networking -ge 2) {
                            foreach ($VlansSubnets in $VmmVlansSubnetsInfo) {
                                Section -Style NOTOCHeading6 -ExcludeFromTOC "$($VlansSubnets.Name)" {
                                    $TableParams = @{
                                        Name = "Subnets & VLANs - $($VlansSubnets.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VlansSubnets | Table @TableParams
                                }
                            }
                        } else {
                            $TableParams = @{
                                Name = "Subnets & VLANs - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'VLAN ID'
                                ColumnWidths = 50, 50
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