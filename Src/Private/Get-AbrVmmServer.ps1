function Get-AbrVmmServer {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft VMM Server information
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
                    Write-PScriboMessage "Collecting VMM Server Settings information."
                    Section -Style Heading2 'VMM Server' {
                        $VmmServerSettingsInfo = @()
                        foreach ($VmmServerSetting in $Vmm) {
                            $InObj = [Ordered]@{
                                'Server FQDN' = $VmmServerSetting.FQDN
                                'IP Address' = (Get-NetIPAddress -CimSession $VMMCimSession -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike "127.0.0.1" }).IPAddress
                                'Product Version' = $VmmServerSetting.ProductVersion
                                'Server Port' = $VmmServerSetting.Port
                                'VM Connect Port' = $VmmServerSetting.VMConnectDefaultPort
                                'VMM Service Account' = $VmmServerSetting.VMMServiceAccount
                                'VMM High Availability' = $VmmServerSetting.IsHighlyAvailable
                            }

                            $VmmServerSettingsInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.Infrastructure -ge 2) {
                            Paragraph "The following sections detail the configuration of the VMM Server Settings."
                            foreach ($VmmServerSetting in $VmmServerSettingsInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($VmmServerSetting.'Server FQDN')" {
                                    $TableParams = @{
                                        Name = "VMM Server Settings - $($VmmServerSetting.'Server FQDN')"
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
                            Paragraph "The following table summarises the configuration of the VMM Server Settings."
                            BlankLine
                            $TableParams = @{
                                Name = "VMM Server Settings - $($VmmServerSetting.FQDN)"
                                List = $false
                                Columns = 'Server FQDN', 'IP Address', 'Product Version', 'Server Port'
                                ColumnWidths = 45, 20, 20, 15
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmServerSettingsInfo | Table @TableParams
                        }
                        Get-AbrVmmRunAsAccount
                        Get-AbrVmmUserRole
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}