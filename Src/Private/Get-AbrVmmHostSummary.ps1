function Get-AbrVmmHostSummary {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Host Summary information
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
        Write-PScriboMessage "Hosts InfoLevel set at $($InfoLevel.Hosts)."
    }

    process {
        try {
            if ($InfoLevel.Hosts -gt 0) {
                if ($VMHosts = Get-SCVMHost -VMMServer $ConnectVmmServer | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Host information."
                    $VmmHostInfo = @()
                    foreach ($VMHost in $VMHosts) {
                        $InObj = [Ordered]@{
                            'Name' = $VMHost.ComputerName
                            'Host Group' = $VMHost.OperatingSystem
                            'Cluster Nodes' = $VMHost.VMHostGroup
                        }

                        $VmmHostInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                    }

                    $TableParams = @{
                        Name = "Host Summary - $($Vmm.FQDN)"
                        List = $false
                        ColumnWidths = 33, 33, 34
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $VmmHostInfo | Table @TableParams
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}