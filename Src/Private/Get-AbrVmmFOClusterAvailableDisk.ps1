function Get-AbrVmmFOClusterAvailableDisk {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Cluster Available Disk
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
        Write-PScriboMessage "Collecting Host Clusters Available Disk information."
    }

    process {
        try {
            $ClusterAvailableDisks = Invoke-Command -Session $ClusterTempPssSession { Get-ClusterAvailableDisk } | Sort-Object -Property Name
            if ($ClusterAvailableDisks) {
                Section -Style Heading3 "Available Disk" {
                    $OutObj = @()
                    foreach ($ClusterAvailableDisk in $ClusterAvailableDisks) {
                        try {
                            $inObj = [ordered] @{
                                'Name' = $ClusterAvailableDisk.Name
                                'Number' = $ClusterAvailableDisk.Number
                                'Size' = ConvertTo-FileSizeString $ClusterAvailableDisk.Size
                            }
                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }

                    $TableParams = @{
                        Name = "Available Disk - $($Cluster)"
                        List = $false
                        ColumnWidths = 40, 30, 30
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