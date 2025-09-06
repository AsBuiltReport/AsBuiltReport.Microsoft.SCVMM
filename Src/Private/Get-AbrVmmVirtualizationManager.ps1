function Get-AbrVmmVirtualizationManager {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Virtualization Manager information
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
                if ($VMMVirtualizationManagers = Get-SCVirtualizationManager -VMMServer $ConnectVmmServer | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Virtualization Managers information."
                    Section -Style Heading3 'Virtualization Managers' {
                        $VmmVirtualizationManagersInfo = @()
                        foreach ($VMMVirtualizationManager in $VMMVirtualizationManagers) {
                            $InObj = [Ordered]@{
                                'Name' = $VMMVirtualizationManager.Name
                                'Version' = $VMMVirtualizationManager.Version
                                'Status' = $VMMVirtualizationManager.Status
                                'Ssl Certificate Hash' = $VMMVirtualizationManager.SslCertificateHash
                                'Number Of Managed Hosts' = $VMMVirtualizationManager.NumberOfManagedHosts
                                'User Name' = $VMMVirtualizationManager.UserName
                                'RunAs Account' = $VMMVirtualizationManager.RunAsAccount
                            }

                            $VmmVirtualizationManagersInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($HealthCheck.Infrastructure) {
                            $VmmVirtualizationManagersInfo | Where-Object { $_.'Status' -ne 'Responding' } | Set-Style -Style Warning -Property 'Status'
                        }

                        if ($InfoLevel.Infrastructure -ge 2) {
                            Paragraph "The following sections detail the configuration of the virtualization managers."
                            foreach ($VMMVirtualizationManager in $VmmVirtualizationManagersInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($VMMVirtualizationManager.Name)" {
                                    $TableParams = @{
                                        Name = "Virtualization Managers - $($VMMVirtualizationManager.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMMVirtualizationManager | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the virtualization managers."
                            BlankLine
                            $TableParams = @{
                                Name = "Virtualization Managers - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'Version', 'RunAs Account', 'Status'
                                ColumnWidths = 35, 25, 20, 20
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmVirtualizationManagersInfo | Table @TableParams
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