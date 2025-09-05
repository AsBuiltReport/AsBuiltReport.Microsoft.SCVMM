function Get-AbrVmmFOClusterPermission {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft Cluster Permissions
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
        Write-PScriboMessage "Collecting Host Cluster Permissions Settings information."
    }

    process {
        try {
            $ClusterAccess = Invoke-Command -Session $ClusterTempPssSession { Get-ClusterAccess } | Sort-Object -Property Identity
            if ($ClusterAccess) {
                Section -Style Heading3 "Access Permissions" {
                    $OutObj = @()
                    foreach ($Permission in $ClusterAccess) {
                        try {
                            $inObj = [ordered] @{
                                'Identity' = $Permission.IdentityReference
                                'Access Control Type' = $Permission.AccessControlType
                                'Rights' = $Permission.ClusterRights
                            }
                            $OutObj += [pscustomobject](ConvertTo-HashToYN $inObj)
                        } catch {
                            Write-PScriboMessage -IsWarning $_.Exception.Message
                        }
                    }

                    $TableParams = @{
                        Name = "Access Permission - $($Cluster)"
                        List = $false
                        ColumnWidths = 60, 20, 20
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