function Get-AbrVmmUplinkPortProfile {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Uplink Port Profile information
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
                if ($UplinkPortProfiles = Get-SCNativeUplinkPortProfile | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Uplink Port Profile information."
                    Section -Style Heading3 'Uplink Port Profiles' {
                        $VmmUplinkPortProfileInfo = @()
                        foreach ($UplinkPortProfile in $UplinkPortProfiles) {
                            $InObj = [Ordered]@{
                                'Name' = $UplinkPortProfile.Name
                                'Team Mode' = $UplinkPortProfile.LBFOTeamMode
                                'Load Balance' = $UplinkPortProfile.LBFOLoadBalancingAlgorithm
                                'Enabled Network Virtualization' = $UplinkPortProfile.EnableNetworkVirtualization
                                'Logical Network Definitions' = $UplinkPortProfile.LogicalNetworkDefinitions -join ', '
                                'Description' = $UplinkPortProfile.Description
                            }

                            $VmmUplinkPortProfileInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.Networking -ge 2) {
                            Paragraph "The following sections detail the configuration of the uplink port profile."
                            foreach ($UplinkPortProfile in $VmmUplinkPortProfileInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($UplinkPortProfile.Name)" {
                                    $TableParams = @{
                                        Name = "Uplink Port Profile - $($UplinkPortProfile.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $UplinkPortProfile | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the uplink port profile."
                            BlankLine
                            $TableParams = @{
                                Name = "Uplink Port Profile - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'Team Mode', 'Load Balance', 'Enabled Network Virtualization'
                                ColumnWidths = 25, 25, 25, 25
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmUplinkPortProfileInfo | Table @TableParams
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