function Get-AbrVmmCluster {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Clusters information
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
                Write-PScriboMessage "Collecting VMM Cluster information."
                if ($ScVmmClusters = Get-SCVMHostCluster) {
                    Section -Style Heading1 'Clusters' {
                        Paragraph "The following table summarises the configuration of the clusters."
                        BlankLine
                        Get-AbrVmmClusterSummary
                        Section -Style Heading3 'Clusters Configuration' {
                            foreach ($ScVmmCluster in ($ScVmmClusters | Where-Object { $_.VirtualizationPlatform -eq 'HyperV' })) {
                                try {
                                    $script:ClusterTempPssSession = New-PSSession $ScVmmCluster.Name -Credential $Credential
                                    $script:ClusterTempCimSession = New-CimSession $ScVmmCluster.Name -Credential $Credential
                                    if ($Cluster = Invoke-Command -Session $script:ClusterTempPssSession { Get-Cluster }) {
                                        Section -Style Heading4 $Cluster.Name {
                                            # Cluster Server Configuration
                                            Get-AbrVmmFOCluster
                                            # Cluster Access Permission
                                            Get-AbrVmmFOClusterPermission
                                            # Cluster Nodes
                                            Get-AbrVmmFOClusterNode
                                            # Cluster Available Disks
                                            Get-AbrVmmFOClusterAvailableDisk
                                            # Cluster Fault Domain
                                            Get-AbrVmmFOClusterFaultDomain
                                            # Cluster Networks
                                            Get-AbrVmmFOClusterNetwork
                                            # Cluster Quorum
                                            Get-AbrVmmFOClusterQuorum
                                            #Cluster Resources
                                            Get-AbrVmmFOClusterResource
                                            #Cluster Shared Volume
                                            Get-AbrVmmFOClusterSharedVolume
                                        }
                                    }
                                } catch {
                                    Write-PScriboMessage -IsWarning $($_.Exception.Message)
                                }
                                if ($ClusterTempPssSession) {
                                    Remove-PSSession $ClusterTempPssSession
                                }
                                if ($ClusterTempCimSession) {
                                    Remove-CimSession $ClusterTempCimSession
                                }
                            }
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