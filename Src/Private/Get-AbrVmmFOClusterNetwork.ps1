function Get-AbrVmmFOClusterNetwork {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Cluster Networks
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
        Write-PScriboMessage "Collecting Host Cluster Networks information."
    }

    process {
        try {
            $ClusterNetworks = Invoke-Command -Session $ClusterTempPssSession { Get-ClusterNetwork } | Sort-Object -Property Name
            if ($ClusterNetworks) {
                Section -Style Heading3 "Networks" {
                    $OutObj = @()
                    foreach ($ClusterNetwork in $ClusterNetworks) {
                        try {
                            $inObj = [ordered] @{
                                'Name' = $ClusterNetwork.Name
                                'State' = $ClusterNetwork.State
                                'Role' = $ClusterNetwork.Role
                                'Address' = "$($ClusterNetwork.Address)/$($ClusterNetwork.AddressMask)"
                            }
                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }


                    if ($HealthCheck.Clusters) {
                        $OutObj | Where-Object { $_.'State' -ne 'UP' } | Set-Style -Style Warning -Property 'State'
                    }

                    $TableParams = @{
                        Name = "Networks - $($Cluster)"
                        List = $false
                        ColumnWidths = 30, 15, 20, 35
                    }
                    if ($Report.ShowTableCaptions) {
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                    }
                    $OutObj | Table @TableParams
                    # Cluster Network Interfaces
                    Get-AbrVmmFOClusterNetworkInterface
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $_.Exception.Message
        }
    }

    end {}

}