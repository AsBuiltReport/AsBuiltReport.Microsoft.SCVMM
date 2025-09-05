function Get-AbrVmmFOClusterNode {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Cluster Nodes
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
    }

    process {
        try {
            $ClusterNodes = Invoke-Command -Session $ClusterTempPssSession { Get-ClusterNode } | Sort-Object -Property Identity
            if ($ClusterNodes) {
                Write-PScriboMessage "Collecting Host Cluster Settings information."
                Section -Style Heading3 'Nodes' {
                    $OutObj = @()
                    foreach ($ClusterNode in $ClusterNodes) {
                        $inObj = [ordered] @{
                            'Name' = $ClusterNode.Name
                            'State' = $ClusterNode.State
                            'Type' = $ClusterNode.Type
                            'Cluster' = $ClusterNode.Cluster
                            'Fault Domain' = $ClusterNode.FaultDomain
                            'Node Weight' = $ClusterNode.NodeWeight
                            'Model' = $ClusterNode.Model
                            'Manufacturer' = $ClusterNode.Manufacturer
                            'Serial Number' = $ClusterNode.SerialNumber
                            'Status Information' = $ClusterNode.StatusInformation
                            'Description' = $ClusterNode.Description

                        }
                        $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                    }

                    if ($HealthCheck.Clusters) {
                        $OutObj | Where-Object { $_.'State' -ne 'UP' } | Set-Style -Style Warning -Property 'State'
                    }

                    if ($InfoLevel.Clusters -ge 2) {
                        Paragraph "The following sections detail the configuration of the Cluster Nodes."
                        foreach ($ClusterNode in $OutObj) {
                            Section -ExcludeFromTOC -Style NOTOCHeading4 "$($ClusterNode.Name)" {
                                $TableParams = @{
                                    Name = "Node - $($ClusterNode.Name)"
                                    List = $true
                                    ColumnWidths = 40, 60
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $ClusterNode | Table @TableParams
                            }
                        }
                    } else {
                        Paragraph "The following table summarizes the configuration of the Cluster Nodes."
                        BlankLine
                        $TableParams = @{
                            Name = "Nodes - $($Cluster)"
                            List = $false
                            Columns = 'Name', 'State', 'Type'
                            ColumnWidths = 40, 30, 30
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Table @TableParams
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning "Cluster Nodes Section: $($_.Exception.Message)"
        }
    }

    end {}
}