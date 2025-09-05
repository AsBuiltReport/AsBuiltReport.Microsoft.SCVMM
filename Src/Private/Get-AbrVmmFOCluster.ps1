function Get-AbrVmmFOCluster {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Cluster configuration
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
        Write-PScriboMessage "Collecting Host Cluster Server information."
    }

    process {
        try {
            $Clusters = Invoke-Command -Session $ClusterTempPssSession { Get-Cluster | Select-Object -Property * }
            if ($Clusters) {
                $OutObj = @()
                try {
                    $inObj = [ordered] @{
                        'Name' = $Clusters.Name
                        'Domain' = $Clusters.Domain
                        'Shared Volumes Root' = $Clusters.SharedVolumesRoot
                        'Administrative Access Point' = $Clusters.AdministrativeAccessPoint
                        'Description' = $Clusters.Description
                    }
                    $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                } catch {
                    Write-PScriboMessage -IsWarning $_.Exception.Message
                }

                $TableParams = @{
                    Name = "Cluster Servers Settings - $($Clusters.Name)"
                    List = $true
                    ColumnWidths = 40, 60
                }
                if ($Report.ShowTableCaptions) {
                    $TableParams['Caption'] = "- $($TableParams.Name)"
                }
                $OutObj | Table @TableParams
            }
        } catch {
            Write-PScriboMessage -IsWarning $_.Exception.Message
        }
    }

    end {}

}