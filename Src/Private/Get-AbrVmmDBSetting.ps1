function Get-AbrVmmDBSetting {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft VMM Database information
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
                    Write-PScriboMessage "Collecting VMM Database Settings information."
                    Section -Style Heading3 'Database Settings' {
                        $VmmServerSettingsInfo = @()
                        foreach ($VmmServerSetting in $Vmm) {
                            $InObj = [Ordered]@{
                                'Server Name' = $VMM.DatabaseServerName
                                'Instance Name' = $VMM.DatabaseInstanceName
                                'DataBase Name' = $VMM.DatabaseName
                                'Version' = $VMM.DatabaseVersion
                            }

                            $VmmServerSettingsInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.Infrastructure -ge 2) {
                            Paragraph "The following sections detail the configuration of the VMM Database Settings."
                            foreach ($VmmServerSetting in $VmmServerSettingsInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($Vmm.FQDN)" {
                                    $TableParams = @{
                                        Name = "Database Settings - $($Vmm.FQDN)"
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
                            Paragraph "The following table summarises the configuration of the VMM Database Settings."
                            BlankLine
                            $TableParams = @{
                                Name = "Database Settings - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Server Name', 'Instance Name', 'DataBase Name', 'Version'
                                ColumnWidths = 30, 30, 20, 20
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