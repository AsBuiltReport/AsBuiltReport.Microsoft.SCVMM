function Get-AbrVmmInfrastructureDiagram {
    <#
    .SYNOPSIS
        Used by As Built Report to built VMM infrastructure diagram
    .DESCRIPTION

    .NOTES
        Version:        0.1.1
        Author:         AsBuiltReport Organization
        Twitter:        @AsBuiltReport
        Github:         AsBuiltReport
    .EXAMPLE

    .LINK

    #>
    [CmdletBinding()]
    param (
    )

    begin {
        Write-PScriboMessage "Generating Infrastructure Diagram for VMM."
        # Used for DraftMode (Don't touch it!)
        if ($Options.EnableDiagramDebug) {
            $EdgeDebug = @{style = 'filled'; color = 'red' }
            $SubGraphDebug = @{style = 'dashed'; color = 'red' }
            $NodeDebug = @{color = 'black'; style = 'red'; shape = 'plain' }
            $NodeDebugEdge = @{color = 'black'; style = 'red'; shape = 'plain' }
            $IconDebug = $true
        } else {
            $EdgeDebug = @{style = 'invis'; color = 'red' }
            $SubGraphDebug = @{style = 'invis'; color = 'gray' }
            $NodeDebug = @{color = 'transparent'; style = 'transparent'; shape = 'point' }
            $NodeDebugEdge = @{color = 'transparent'; style = 'transparent'; shape = 'none' }
            $IconDebug = $false
        }

        # Used for setting diagram Theme (Can be change to fits your needs!)
        if ($Options.DiagramTheme -eq 'Black') {
            $Edgecolor = 'White'
            $Fontcolor = 'White'
        } elseif ($Options.DiagramTheme -eq 'Neon') {
            $Edgecolor = 'gold2'
            $Fontcolor = 'gold2'
        } else {
            $Edgecolor = '#71797E'
            $Fontcolor = '#565656'
        }
    }

    process {
        try {
            if ($VMM) {
                $UpdateServer = Get-SCUpdateServer -VMMServer $ConnectVmmServer
                $LBRServers = Get-SCLibraryServer -VMMServer $ConnectVmmServer | Sort-Object -Property Name
                $VMMVirtualizationManagers = Get-SCVirtualizationManager -VMMServer $ConnectVmmServer | Sort-Object -Property Name

                $VMMServerAdditionalInfo = [pscustomobject][Ordered]@{
                    'IP Address' = Get-NodeIP -Hostname $VMM.FQDN
                    'Server Port' = $VMM.Port
                    'Version' = $VMM.ProductVersion
                    'Role' = "VMM Server"

                }
                $VMMDBServerAdditionalInfo = [pscustomobject][Ordered]@{
                    'Server' = $VMM.DatabaseServerName
                    'Instance' = $VMM.DatabaseInstanceName
                    'Server Port' = '1431'
                    'Version' = $VMM.DatabaseVersion
                }

                Node VMMserver @{Label = Add-DiaNodeIcon -Name $VMM.FQDN.split(".")[0].ToUpper() -AditionalInfo $VMMServerAdditionalInfo -ImagesObj $Images -IconType "Server" -Align "Center" -IconDebug $IconDebug -FontSize 18; shape = 'plain'; fillColor = 'transparent'; fontsize = 14 }

                Node DBServer @{Label = Add-DiaNodeIcon -Name $VMM.DatabaseName -AditionalInfo $VMMDBServerAdditionalInfo -ImagesObj $Images -IconType "DB_Server" -Align "Center" -IconDebug $IconDebug -FontSize 18; shape = 'plain'; fillColor = 'transparent'; fontsize = 14 }

                Rank VMMserver, DBServer

                Edge -From VMMserver -To DBServer -Attributes @{minlen = 2; label = "DB Connection 1431"; color = $Edgecolor; fontcolor = $Fontcolor; fontsize = 16; style = 'dashed'; penwidth = 2; arrowhead = 'normal'; arrowtail = 'dot' }

                if ($UpdateServer) {
                    $VMMUDServerAdditionalInfo = [pscustomobject][Ordered]@{
                        'IP Address' = Get-NodeIP -Hostname $UpdateServer.Name
                        'Server Port' = $UpdateServer.Port
                        'Type' = $UpdateServer.ServerType
                        'Role' = "Update Server"

                    }

                    Node UpdateServer @{Label = Add-DiaNodeIcon -Name $UpdateServer.Name.split(".")[0].ToUpper() -AditionalInfo $VMMUDServerAdditionalInfo -ImagesObj $Images -IconType "Server" -Align "Center" -IconDebug $IconDebug -FontSize 18; shape = 'plain'; fillColor = 'transparent'; fontsize = 14 }
                } else {
                    Write-PScriboMessage "Unable to create UpdateServer Node. No UpdateServer Object found."
                }

                if ($LBRServers) {
                    $LBRAdditionalInfo = @()
                    foreach ($LBR in $LBRServers) {
                        $LBRAdditionalInfo += [pscustomobject][Ordered]@{
                            'IP Address' = Get-NodeIP -Hostname $LBR.Name
                            'Status' = $LBR.Status
                            'Library Group' = switch ([string]::IsNullOrEmpty($LBR.LibraryGroup)) {
                                $true { 'None' }
                                $false { $LBR.LibraryGroup }
                                Default { 'Unknown' }
                            }
                            'Domain' = $LBR.DomainName
                        }
                    }

                    if ($LBRServers.Count -eq 1) {
                        $LBRServersColumnSize = 1
                    } elseif ($ColumnSize) {
                        $LBRServersColumnSize = $ColumnSize
                    } else {
                        $LBRServersColumnSize = $LBRServers.Count
                    }

                    $LBRServersNodeObj = Add-DiaHTMLNodeTable -ImagesObj $Images -inputObject ($LBRServers | ForEach-Object { $_.Name.Split('.')[0] }).ToUpper() -Align "Center" -iconType "Server" -columnSize $LBRServersColumnSize -IconDebug $IconDebug -MultiIcon -AditionalInfo $LBRAdditionalInfo -Subgraph -SubgraphIconType "Server" -SubgraphLabel "Library Servers" -SubgraphLabelPos "top" -SubgraphTableStyle "dashed,rounded" -TableBorderColor "#71797E" -TableBorder "1" -SubgraphLabelFontsize 22 -fontSize 18
                }

                if ($LBRServersNodeObj) {
                    Node LibraryServers @{Label = $LBRServersNodeObj; shape = 'plain'; fillColor = 'transparent'; fontsize = 14 }
                } else {
                    Write-PScriboMessage "Unable to create LibraryServers Node. No LibraryServers Object found."
                }

                if ($VMMVirtualizationManagers) {
                    $VManagerAdditionalInfo = @()
                    foreach ($VManager in $VMMVirtualizationManagers) {
                        $VManagerAdditionalInfo += [pscustomobject][Ordered]@{
                            'IP Address' = Get-NodeIP -Hostname $VManager.Name
                            'Status' = $VManager.Status
                            'Managed Hosts' = $VManager.NumberOfManagedHosts
                            'Managed VMs' = $VManager.NumberOfManagedVirtualMachines
                            'Version' = $VManager.Version
                        }
                    }

                    if ($VMMVirtualizationManagers.Count -eq 1) {
                        $VManagerColumnSize = 1
                    } elseif ($ColumnSize) {
                        $VManagerColumnSize = $ColumnSize
                    } else {
                        $VManagerColumnSize = $VMMVirtualizationManagers.Count
                    }

                    $VManagerNodeObj = Add-DiaHTMLNodeTable -ImagesObj $Images -inputObject $VMMVirtualizationManagers.Name.toUpper() -Align "Center" -iconType "Server" -columnSize $VManagerColumnSize -IconDebug $IconDebug -MultiIcon -AditionalInfo $VManagerAdditionalInfo -Subgraph -SubgraphIconType "Server" -SubgraphLabel "VMware vCenters" -SubgraphLabelPos "top" -SubgraphTableStyle "dashed,rounded" -TableBorderColor "#71797E" -TableBorder "1" -SubgraphLabelFontsize 22 -fontSize 18
                }

                if ($VManagerNodeObj) {
                    Node VirtualizationManagers @{Label = $VManagerNodeObj; shape = 'plain'; fillColor = 'transparent'; fontsize = 14 }
                } else {
                    Write-PScriboMessage "Unable to create VirtualizationManagers Node. No VirtualizationManagers Object found."
                }

                # VMM elements point of connection (Dummy Nodes!)
                $Node = @('VMMLibraryServersPoint', 'VMMServerPointSpace')
                $NodeEdge = @()

                $LastPoint = 'VMMUpdateServerPointSpace'

                if ($UpdateServer) {
                    $Node += 'VMMUpdateServerPointSpace'
                    $NodeEdge += 'VMMUpdateServerPointSpace'
                    $LastPoint = 'VMMUpdateServerPointSpace'
                }

                if ($VManagerNodeObj) {
                    $Node += 'VMMVirtualizationManagers'
                    $NodeEdge += 'VMMVirtualizationManagers'
                    $LastPoint = 'VMMVirtualizationManagers'
                }

                Node $Node -NodeScript { $_ } @{Label = { $_ } ; fontcolor = $NodeDebug.color; fillColor = $NodeDebug.style; shape = $NodeDebug.shape }

                $NodeStartEnd = @('VMMStartPoint', 'VMMEndPointSpace')
                Node $NodeStartEnd -NodeScript { $_ } @{Label = { $_ }; fillColor = $Edgecolor; fontcolor = $NodeDebug.color; shape = 'point'; fixedsize = 'true'; width = .2 ; height = .2 }
                #---------------------------------------------------------------------------------------------#
                #                             Graphviz Rank Section                                           #
                #                     Rank allow to put Nodes on the same group level                         #
                #         PSgraph: https://psgraph.readthedocs.io/en/stable/Command-Rank-Advanced/            #
                #                     Graphviz: https://graphviz.org/docs/attrs/rank/                         #
                #---------------------------------------------------------------------------------------------#

                # Put the dummy node in the same rank to be able to create a horizontal line
                Rank $NodeStartEnd, $Node

                #---------------------------------------------------------------------------------------------#
                #                             Graphviz Edge Section                                           #
                #                   Edges are Graphviz elements use to interconnect Nodes                     #
                #                 Edges can have attribues like Shape, Size, Styles etc..                     #
                #              PSgraph: https://psgraph.readthedocs.io/en/latest/Command-Edge/                #
                #                      Graphviz: https://graphviz.org/docs/edges/                             #
                #---------------------------------------------------------------------------------------------#

                # LastPoint Min length
                $LastPointMinLen = 5
                # Connect the Dummy Node in a straight line
                # VMMStartPoint --- VMMLibraryServersPoint --- VMMServerPointSpace --- VMMUpdateServerPointSpace
                Edge -From VMMStartPoint -To VMMLibraryServersPoint @{minlen = 5; arrowtail = 'none'; arrowhead = 'none'; style = 'filled' }
                Edge -From VMMLibraryServersPoint -To VMMServerPointSpace @{minlen = 5; arrowtail = 'none'; arrowhead = 'none'; style = 'filled' }
                if ($UpdateServer) {
                    Edge -From VMMServerPointSpace -To VMMUpdateServerPointSpace @{minlen = 5; arrowtail = 'none'; arrowhead = 'none'; style = 'filled' }
                } elseif ($VManagerNodeObj) {
                    Edge -From VMMServerPointSpace -To $LastPoint @{minlen = 5; arrowtail = 'none'; arrowhead = 'none'; style = 'filled' }
                } else {
                    $LastPoint = 'VMMServerPointSpace'
                    $LastPointMinLen = 8
                }

                if ($UpdateServer) {
                    Edge -From VMMUpdateServerPointSpace -To UpdateServer @{minlen = 5; arrowtail = 'none'; arrowhead = 'dot'; style = 'filled' }
                }

                if ($LBRServers) {
                    Edge -From VMMLibraryServersPoint -To LibraryServers @{minlen = 5; arrowtail = 'none'; arrowhead = 'dot'; style = 'filled' }
                }

                if ($VManagerNodeObj) {
                    Edge -From VMMVirtualizationManagers -To VirtualizationManagers @{minlen = 5; arrowtail = 'none'; arrowhead = 'dot'; style = 'filled' }
                }

                # Connect the available Points
                $index = 0
                foreach ($Element in $NodeEdge) {
                    $index++
                    Edge -From $Element -To $NodeEdge[$index] @{minlen = 7; arrowtail = 'none'; arrowhead = 'none'; style = 'filled' }
                }

                ####################################################################################
                #                                                                                  #
                #      This section connect the Infrastructure component to the Dummy Points       #
                #                                                                                  #
                ####################################################################################

                # Connect VMM server to the Dummy line
                Edge -From VMMserver -To VMMServerPointSpace @{minlen = 2; arrowtail = 'dot'; arrowhead = 'none'; style = 'dashed' }


                ####################################################################################
                #                                                                                  #
                #   This section connect the Last Infrastructure component to VMMEndPointSpace     #
                #                                                                                  #
                ####################################################################################

                if ($LastPoint) {
                    Edge -From $LastPoint -To VMMEndPointSpace @{minlen = $LastPointMinLen; arrowtail = 'none'; arrowhead = 'none'; style = 'filled' }
                }

            }
        } catch {
            Write-PScriboMessage -IsWarning $_.Exception.Message
        }
    }

    end {}

}