function Get-AbrVmmFOClusterResource {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Cluster Resource
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
        Write-PScriboMessage "Collecting Host Cluster Resource information."
    }

    process {
        try {
            $ClusterResources = Invoke-Command -Session $ClusterTempPssSession { Get-ClusterResource | Select-Object -Property * } | Sort-Object -Property Name
            if ($ClusterResources) {
                Section -Style Heading3 "Resource" {
                    $OutObj = @()
                    foreach ($ClusterResource in $ClusterResources) {
                        try {
                            $inObj = [ordered] @{
                                'Name' = $ClusterResource.Name
                                'Owner Group' = $ClusterResource.OwnerGroup
                                'Resource Type' = $ClusterResource.ResourceType
                                'State' = $ClusterResource.State
                            }
                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }


                    if ($HealthCheck.Clusters) {
                        $OutObj | Where-Object { $_.'State' -notlike 'Online' } | Set-Style -Style Warning -Property 'State'
                    }

                    $TableParams = @{
                        Name = "Resource - $($Cluster)"
                        List = $false
                        ColumnWidths = 25, 25, 35, 15
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