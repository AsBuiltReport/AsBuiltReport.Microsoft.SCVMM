function Get-AbrVmmServerConnection {
    <#
    .SYNOPSIS
    Used by As Built Report to establish connection to Microsoft SCVMM Server.
    .DESCRIPTION
        Documents the configuration of Microsoft SCVMM in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.1.1
        Author:         Jonathan Colon
        Twitter:        @jcolonfzenpr
        Github:         rebelinux
        Credits:        Iain Brighton (@iainbrighton) - PScribo module

    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM
    #>
    [CmdletBinding()]
    param (

    )

    begin {
        Write-PScriboMessage "Establishing initial connection to VMM Server: $($System)."
    }

    process {
        Write-PScriboMessage "Looking for VMM existing server connection."
        $script:ConnectVmmServer = Get-SCVMMServer -ComputerName $Server -Credential $Credential -TCPPort $Options.VmmServerPort
        if ($ConnectVmmServer) {
            Write-PScriboMessage "Successfully connected to $($System):$($Options.VmmServerPort) Vmm Server."
            $script:VMM = $ConnectVmmServer
            $script:VMMCimSession = New-CimSession -ComputerName ($VMM.FQDN) -Credential $Credential
        } else {
            Write-PScriboMessage -IsWarning $_.Exception.Message
            throw "Failed to connect to Vmm Server Host $($System):$($Options.VmmServerPort) with username $($Credential.USERNAME)"
        }
    }
    end {}
}