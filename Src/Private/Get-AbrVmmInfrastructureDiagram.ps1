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
                $UpdateServer = Get-SCUpdateServer
                $VMMServerAdditionalInfo = [pscustomobject][Ordered]@{
                    'IP Address' = (Get-NetIPAddress -CimSession $VMMCimSession -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike "127.0.0.1" })[0].IPAddress
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
                $VMMUDServerAdditionalInfo = [pscustomobject][Ordered]@{
                    'IP Address' = Get-NodeIP -Hostname $UpdateServer.Name
                    'Server Port' = $UpdateServer.Port
                    'Type' = $UpdateServer.ServerType
                    'Role' = "Update Server"

                }

                Node VMMserver @{Label = Add-DiaNodeIcon -Name $VMM.FQDN.split(".")[0].ToUpper() -AditionalInfo $VMMServerAdditionalInfo -ImagesObj $Images -IconType "Server" -Align "Center" -IconDebug $IconDebug -FontSize 18; shape = 'plain'; fillColor = 'transparent'; fontsize = 14 }

                Node DBServer @{Label = Add-DiaNodeIcon -Name $VMM.DatabaseName -AditionalInfo $VMMDBServerAdditionalInfo -ImagesObj $Images -IconType "DB_Server" -Align "Center" -IconDebug $IconDebug -FontSize 18; shape = 'plain'; fillColor = 'transparent'; fontsize = 14 }

                Node UpdateServer @{Label = Add-DiaNodeIcon -Name $UpdateServer.Name.split(".")[0].ToUpper() -AditionalInfo $VMMUDServerAdditionalInfo -ImagesObj $Images -IconType "Server" -Align "Center" -IconDebug $IconDebug -FontSize 18; shape = 'plain'; fillColor = 'transparent'; fontsize = 14 }

                Edge -From VMMserver -To DBServer -Attributes @{minlen = 2; label = "DB Connection 1431"; color = $Edgecolor; fontcolor = $Fontcolor; fontsize = 16; style = 'dashed'; penwidth = 2; arrowhead = 'normal'; arrowtail = 'none' }

                Edge -From VMMserver -To UpdateServer -Attributes @{minlen = 2; color = $Edgecolor; fontcolor = $Fontcolor; fontsize = 16; style = 'dashed'; penwidth = 2; arrowhead = 'normal'; arrowtail = 'none' }

                Rank VMMserver, DBServer
            }
        } catch {
            Write-PScriboMessage -IsWarning $_.Exception.Message
        }
    }

    end {}

}