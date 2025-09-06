function Get-AbrVmmVmVlanSubnet {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM VM VLANs and Subnets information
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
        $SubnetVLans,
        [string] $Name
    )

    begin {
        Write-PScriboMessage "Networking InfoLevel set at $($InfoLevel.Networking)."
    }

    process {
        try {
            if ($InfoLevel.Networking -gt 0) {
                if ($SubnetVLans) {
                    Write-PScriboMessage "Collecting VMM VM VLANs and Subnets information."
                    Section -Style NOTOCHeading5 -ExcludeFromTOC 'Subnets & VLANs' {
                        $VmmVlansSubnetsInfo = @()
                        foreach ($VlansSubnets in ($SubnetVLans | Sort-Object -Property Name)) {
                            $InObj = [Ordered]@{
                                'Name' = $VlansSubnets.Subnet
                                'VLAN ID' = $VlansSubnets.VLanID
                                'Supports DHCP' = $VlansSubnets.SupportsDHCP
                                'Is Assigned To VM Subnet?' = $VlansSubnets.IsAssignedToVMSubnet
                                'Enabled' = $VlansSubnets.IsVLanEnabled
                            }

                            $VmmVlansSubnetsInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }
                        BlankLine
                        $TableParams = @{
                            Name = "Subnets & VLANs - $($Name)"
                            List = $false
                            ColumnWidths = 20, 20, 20, 20, 20
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $VmmVlansSubnetsInfo | Table @TableParams
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}