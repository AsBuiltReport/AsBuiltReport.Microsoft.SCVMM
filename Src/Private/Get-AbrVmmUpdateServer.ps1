function Get-AbrVmmUpdateServer {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Update Servers information
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
                if ($VMMUpdateServers = Get-SCUpdateServer | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Update Servers information."
                    Section -Style Heading3 'Update Servers' {
                        $VmmUpdateServersInfo = @()
                        foreach ($VMMUpdateServer in $VMMUpdateServers) {
                            $InObj = [Ordered]@{
                                'Name' = $VMMUpdateServer.Name
                                'Port' = $VMMUpdateServer.Port
                                'Is Connection Secure' = $VMMUpdateServer.IsConnectionSecure
                                'Server Type' = $VMMUpdateServer.ServerType
                                'Version' = $VMMUpdateServer.Version
                                'Synchronization Type' = $VMMUpdateServer.SynchronizationType
                                'UsesProxy' = $VMMUpdateServer.UsesProxy
                            }

                            $VmmUpdateServersInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.Infrastructure -ge 2) {
                            Paragraph "The following sections detail the configuration of the update server."
                            foreach ($VMMUpdateServer in $VmmUpdateServersInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($VMMUpdateServer.Name)" {
                                    $TableParams = @{
                                        Name = "Update Servers - $($VMMUpdateServer.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMMUpdateServer | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the update servers."
                            BlankLine
                            $TableParams = @{
                                Name = "Update Servers - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'Port', 'Is Connection Secure', 'Server Type'
                                ColumnWidths = 45, 15, 20, 20
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmUpdateServersInfo | Table @TableParams
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