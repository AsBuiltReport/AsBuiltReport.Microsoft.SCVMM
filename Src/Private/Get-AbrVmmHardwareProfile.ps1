function Get-AbrVmmHardwareProfile {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Hardware Profile information
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
        Write-PScriboMessage "LibraryTemplates InfoLevel set at $($InfoLevel.LibraryTemplates)."
    }

    process {
        try {
            if ($InfoLevel.LibraryTemplates -gt 0) {
                if ($HardwareProfiles = Get-SCHardwareProfile | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Hardware Profile information."
                    Section -Style Heading3 'Hardware Profiles' {
                        $VmmHardwareProfilesInfo = @()
                        foreach ($HardwareProfile in $HardwareProfiles) {
                            $InObj = [Ordered]@{
                                'Name' = $HardwareProfile.Name
                                'CPU Count' = $HardwareProfile.CPUCount
                                'Memory MB' = $HardwareProfile.Memory
                                'Secure Boot Enabled' = $HardwareProfile.SecureBootEnabled
                                'Generation' = $HardwareProfile.Generation
                                'Checkpoint Type' = $HardwareProfile.CheckpointType
                                'Capability Profile' = $HardwareProfile.CapabilityProfile
                                'Monitor Maximum Count' = $HardwareProfile.MonitorMaximumCount
                            }

                            $VmmHardwareProfilesInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.LibraryTemplates -ge 2) {
                            Paragraph "The following sections detail the configuration of the hardware profiles."
                            foreach ($HardwareProfile in $VmmHardwareProfilesInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($HardwareProfile.Name)" {
                                    $TableParams = @{
                                        Name = "Hardware Profiles - $($HardwareProfile.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $HardwareProfile | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the hardware profiles."
                            BlankLine
                            $TableParams = @{
                                Name = "Hardware Profiles - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'CPU Count', 'Memory MB', 'Generation'
                                ColumnWidths = 25, 25, 25, 25
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmHardwareProfilesInfo | Table @TableParams
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