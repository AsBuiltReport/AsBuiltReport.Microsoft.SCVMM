function Get-AbrVmmFOClusterSharedVolumeState {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Cluster Shared Volume State
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
        Write-PScriboMessage "Collecting Host Cluster Shared Volume State information."
    }

    process {
        try {
            $ClusterSharedVolumeStates = Invoke-Command -Session $ClusterTempPssSession { Get-ClusterSharedVolumeState | Select-Object -Property * } | Sort-Object -Property Name
            if ($ClusterSharedVolumeStates) {
                Section -Style Heading4 "Cluster Shared Volume State" {
                    $OutObj = @()
                    foreach ($ClusterSharedVolumeState in $ClusterSharedVolumeStates) {
                        try {
                            $inObj = [ordered] @{
                                'Name' = $ClusterSharedVolumeState.Name
                                'Node' = $ClusterSharedVolumeState.Node
                                'State' = $ClusterSharedVolumeState.StateInfo
                                'Volume Name' = $ClusterSharedVolumeState.VolumeFriendlyName
                                'Volume Path' = $ClusterSharedVolumeState.VolumeName
                            }
                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }

                    if ($HealthCheck.Clusters) {
                        $OutObj | Where-Object { $_.State.Value -eq 'Unavailable' } | Set-Style -Style Warning -Property 'State'
                    }

                    $TableParams = @{
                        Name = "Cluster Shared Volume State - $($Cluster)"
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