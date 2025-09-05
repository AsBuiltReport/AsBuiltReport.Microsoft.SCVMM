function Get-AbrVmmFOClusterSharedVolume {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Cluster Shared Volume
    .DESCRIPTION
        Documents the configuration of Microsoft Windows Server in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.5.5
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
        Write-PScriboMessage "Collecting Host Cluster Shared Volume information."
    }

    process {
        try {
            # Todo: Add more properties if needed
            $ClusterSharedVolumes = Invoke-Command -Session $ClusterTempPssSession { Get-ClusterSharedVolume | Select-Object -Property * } | Sort-Object -Property Name
            if ($ClusterSharedVolumes) {
                Section -Style Heading3 "Cluster Shared Volume" {
                    $OutObj = @()
                    foreach ($ClusterSharedVolume in $ClusterSharedVolumes) {
                        try {
                            $inObj = [ordered] @{
                                'Name' = $ClusterSharedVolume.Name
                                'Owner Node' = $ClusterSharedVolume.OwnerNode
                                'Shared Volume' = switch ([string]::IsNullOrEmpty($ClusterSharedVolume.SharedVolumeInfo.FriendlyVolumeName)) {
                                    $true { "Unknown" }
                                    $false { $ClusterSharedVolume.SharedVolumeInfo.FriendlyVolumeName }
                                    default { "--" }
                                }
                                'File System Type' = $ClusterSharedVolume.SharedVolumeInfo.Partition.FileSystem
                                'Volume Capacity(GB)' = [Math]::Round(($ClusterSharedVolume.SharedVolumeInfo.Partition.Size) / 1gb)
                                'Free Capacity(GB)' = [Math]::Round(($ClusterSharedVolume.SharedVolumeInfo.Partition.FreeSpace) / 1gb)
                                'State' = $ClusterSharedVolume.State
                            }
                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }


                    if ($HealthCheck.Clusters) {
                        $OutObj | Where-Object { $_.'State' -notlike 'Online' } | Set-Style -Style Warning -Property 'State'
                    }

                    if ($InfoLevel.Clusters -ge 2) {
                        Paragraph "The following sections detail the configuration of the Cluster Shared Volume."
                        foreach ($ClusterSharedVolume in $OutObj) {
                            Section -ExcludeFromTOC -Style NOTOCHeading4 "$($ClusterSharedVolume.Name)" {
                                $TableParams = @{
                                    Name = "Cluster Shared Volume - $($ClusterSharedVolume.Name)"
                                    List = $true
                                    ColumnWidths = 40, 60
                                }
                                if ($Report.ShowTableCaptions) {
                                    $TableParams['Caption'] = "- $($TableParams.Name)"
                                }
                                $ClusterSharedVolume | Table @TableParams
                            }
                        }
                    } else {
                        Paragraph "The following table summarizes the configuration of the Cluster Shared Volume."
                        BlankLine
                        $TableParams = @{
                            Name = "Cluster Shared Volumes - $($Cluster)"
                            List = $false
                            Columns = 'Name', 'Owner Node', 'Shared Volume', 'State'
                            ColumnWidths = 25, 25, 35, 15
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $OutObj | Table @TableParams
                    }
                    #Cluster Shared Volume State
                    Get-AbrVmmFOClusterSharedVolumeState
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $_.Exception.Message
        }
    }

    end {}

}