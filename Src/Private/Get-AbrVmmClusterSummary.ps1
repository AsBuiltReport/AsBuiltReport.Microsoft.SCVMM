function Get-AbrVmmClusterSummary {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Cluster Summary information
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
        Write-PScriboMessage "Clusters InfoLevel set at $($InfoLevel.Clusters)."
    }

    process {
        try {
            if ($InfoLevel.Clusters -gt 0) {
                if ($ScVmmClusters = Get-SCVMHostCluster -VMMServer $ConnectVmmServer | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Cluster information."
                    $VmmClusterInfo = @()
                    foreach ($ScVmmCluster in $ScVmmClusters) {
                        $InObj = [Ordered]@{
                            'Name' = $ScVmmCluster.Name
                            'Host Group' = $ScVmmCluster.HostGroup
                            'Cluster Nodes' = $ScVmmCluster.Nodes
                        }

                        $VmmClusterInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                    }

                    $TableParams = @{
                        Name = "Cluster Summary - $($Vmm.FQDN)"
                        List = $false
                        ColumnWidths = 33, 33, 34
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $VmmClusterInfo | Table @TableParams
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}