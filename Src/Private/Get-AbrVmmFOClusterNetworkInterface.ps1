function Get-AbrVmmFOClusterNetworkInterface {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Cluster Network Interfaces
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
        Write-PScriboMessage "Collecting Host Cluster Network Interface information."
    }

    process {
        try {
            $ClusterNetworkInterfaces = Invoke-Command -Session $ClusterTempPssSession { Get-ClusterNetworkInterface } | Sort-Object -Property Name
            if ($ClusterNetworkInterfaces) {
                Section -Style Heading3 "Interfaces" {
                    $OutObj = @()
                    foreach ($ClusterNetworkInterface in $ClusterNetworkInterfaces) {
                        try {
                            $inObj = [ordered] @{
                                'Name' = $ClusterNetworkInterface.Name
                                'Node' = $ClusterNetworkInterface.Node
                                'Network' = $ClusterNetworkInterface.Network
                                'State' = $ClusterNetworkInterface.State
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
                        Name = "Interfaces - $($Cluster)"
                        List = $false
                        ColumnWidths = 30, 25, 30, 15
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