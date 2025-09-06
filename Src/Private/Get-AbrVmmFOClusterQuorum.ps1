function Get-AbrVmmFOClusterQuorum {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Cluster Quorum
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.2
        Author:         Jonathan Colon
        Twitter:        @jcolonfzenpr
        Github:         rebelinux
        Credits:        Iain Brighton (@iainbrighton) - PScribo module

    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.Windows
    #>

    [CmdletBinding()]
    param (
    )

    begin {
        Write-PScriboMessage "Clusters InfoLevel set at $($InfoLevel.Clusters)."
        Write-PScriboMessage "Collecting Host Cluster Quorum information."
    }

    process {
        try {
            $ClusterQuorums = Invoke-Command -Session $ClusterTempPssSession { Get-ClusterQuorum | Select-Object -Property * } | Sort-Object -Property Name
            if ($ClusterQuorums) {
                Section -Style Heading3 "Quorum" {
                    $OutObj = @()
                    foreach ($ClusterQuorum in $ClusterQuorums) {
                        try {
                            $inObj = [ordered] @{
                                'Name' = $ClusterQuorum.QuorumResource.Name
                                'In State' = $ClusterQuorum.QuorumType
                                'Status' = $ClusterQuorum.QuorumResource.State
                                'Owner Node' = $ClusterQuorum.QuorumResource.OwnerNode
                                'Resource Type' = $ClusterQuorum.QuorumResource.ResourceType
                            }
                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }

                    $TableParams = @{
                        Name = "Quorum - $($Cluster)"
                        List = $false
                        ColumnWidths = 20, 20, 20, 20, 20
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $OutObj | Table @TableParams
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $_.Exception.Message
        }
    }

    end {}

}