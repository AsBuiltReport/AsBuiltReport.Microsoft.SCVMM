function Get-AbrVmmNetworkAdapterPortProfile {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Network Adapter Port Profiles information
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
                if ($VirtualNetworkAdapterPortProfiles = Get-SCVirtualNetworkAdapterNativePortProfile | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Network Adapter Port Profiles information."
                    Section -Style Heading3 'Network Adapter Port Profiles' {
                        $VmmNetworkAdapterPortProfileInfo = @()
                        foreach ($VirtualNetworkAdapterPortProfile in $VirtualNetworkAdapterPortProfiles) {
                            $InObj = [Ordered]@{
                                'Name' = $VirtualNetworkAdapterPortProfile.Name
                                'Teaming' = $VirtualNetworkAdapterPortProfile.AllowTeaming
                                'Mac Address Spoofing' = $VirtualNetworkAdapterPortProfile.AllowMacAddressSpoofing
                                'Ieee Priority Tagging' = $VirtualNetworkAdapterPortProfile.AllowIeeePriorityTagging
                                'DHCP Guard' = $VirtualNetworkAdapterPortProfile.EnableDHCPGuard
                                'Guest IP Network Virtualization Updates' = $VirtualNetworkAdapterPortProfile.EnableGuestIPNetworkVirtualizationUpdates
                                'Router Guard' = $VirtualNetworkAdapterPortProfile.EnableRouterGuard
                                'Minimum Bandwidth Weight' = $VirtualNetworkAdapterPortProfile.MinimumBandwidthWeight
                                'Minimum Bandwidth Absolute In Mbps' = $VirtualNetworkAdapterPortProfile.MinimumBandwidthAbsoluteInMbps
                                'Maximum Bandwidth Absolute In Mbps' = $VirtualNetworkAdapterPortProfile.MaximumBandwidthAbsoluteInMbps
                                'Enable Vmq' = $VirtualNetworkAdapterPortProfile.EnableVmq
                                'Enable IPsec Offload' = $VirtualNetworkAdapterPortProfile.EnableIPsecOffload
                                'Enable Iov' = $VirtualNetworkAdapterPortProfile.EnableIov
                                'Enable Vrss' = $VirtualNetworkAdapterPortProfile.EnableVrss
                                'Enable Rdma' = $VirtualNetworkAdapterPortProfile.EnableRdma
                            }

                            $VmmNetworkAdapterPortProfileInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.Networking -ge 2) {
                            Paragraph "The following sections detail the configuration of the network adapter port profiles."
                            foreach ($VirtualNetworkAdapterPortProfile in $VmmNetworkAdapterPortProfileInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($VirtualNetworkAdapterPortProfile.Name)" {
                                    $TableParams = @{
                                        Name = "Network Adapter Port Profiles - $($VirtualNetworkAdapterPortProfile.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VirtualNetworkAdapterPortProfile | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the network adapter port profiles."
                            BlankLine
                            $TableParams = @{
                                Name = "Network Adapter Port Profiles - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'Teaming', 'Mac Address Spoofing', 'DHCP Guard', 'Router Guard'
                                ColumnWidths = 20, 20, 20, 20, 20
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmNetworkAdapterPortProfileInfo | Table @TableParams
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