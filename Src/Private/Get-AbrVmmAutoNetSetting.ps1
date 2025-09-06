function Get-AbrVmmAutoNetSetting {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft VMM Auto Network information
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
        Write-PScriboMessage "Infrastructure InfoLevel set at $($InfoLevel.Infrastructure)."
    }

    process {
        try {
            if ($InfoLevel.Infrastructure -gt 0) {
                if ($Vmm) {
                    Write-PScriboMessage "Collecting VMM Auto Network information."
                    Section -Style Heading3 'AutoNetwork Settings' {
                        $VmmServerSettingsInfo = @()
                        foreach ($VmmServerSetting in $Vmm) {
                            $InObj = [Ordered]@{
                                'Logical Network Creation Enabled' = $VMM.AutomaticLogicalNetworkCreationEnabled
                                'Virtual Network Creation Enabled' = $VMM.AutomaticVirtualNetworkCreationEnabled
                                'Logical Network Match' = $VMM.LogicalNetworkMatchOption
                                'Backup Network Match' = $VMM.BackupLogicalNetworkMatchOption
                            }

                            $VmmServerSettingsInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.Infrastructure -ge 2) {
                            Paragraph "The following sections detail the configuration of the VMM Auto Network Settings."
                            foreach ($VmmServerSetting in $VmmServerSettingsInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($Vmm.FQDN)" {
                                    $TableParams = @{
                                        Name = "Auto Network Settings - $($Vmm.FQDN)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VmmServerSetting | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the VMM Auto Network Settings."
                            BlankLine
                            $TableParams = @{
                                Name = "Auto Network Settings - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Logical Network Creation Enabled', 'Virtual Network Creation Enabled', 'Logical Network Match', 'Backup Network Match'
                                ColumnWidths = 25, 25, 25, 25
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmServerSettingsInfo | Table @TableParams
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