function Get-AbrVmmFOClusterFaultDomain {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Cluster Fault Domain
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
        Write-PScriboMessage "Collecting Host Clusters Fault Domain information."
    }

    process {
        try {
            $ClusterFaultDomains = Get-ClusterFaultDomain -CimSession $ClusterTempCimSession | Sort-Object -Property Name
            if ($ClusterFaultDomains) {
                Section -Style Heading3 "Fault Domain" {
                    $OutObj = @()
                    foreach ($ClusterFaultDomain in $ClusterFaultDomains) {
                        try {
                            $inObj = [ordered] @{
                                'Name' = $ClusterFaultDomain.Name
                                'Type' = $ClusterFaultDomain.Type
                                'Parent Name' = Switch ([string]::IsNullOrEmpty($ClusterFaultDomain.ParentName)) {
                                    $true { "--" }
                                    $false { $ClusterFaultDomain.ParentName }
                                    default { 'Unknown' }
                                }
                                'Children Names' = Switch ([string]::IsNullOrEmpty($ClusterFaultDomain.ChildrenNames)) {
                                    $true { "--" }
                                    $false { $ClusterFaultDomain.ChildrenNames }
                                    default { 'Unknown' }
                                }
                                'Location' = $ClusterFaultDomain.Location
                            }
                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }

                    $TableParams = @{
                        Name = "Fault Domain - $($Cluster)"
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