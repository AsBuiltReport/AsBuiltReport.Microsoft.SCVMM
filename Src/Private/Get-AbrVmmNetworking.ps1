function Get-AbrVmmNetworking {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Networking information
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
        Write-PScriboMessage "Networking InfoLevel set at $($InfoLevel.Networking)."
    }

    process {
        try {
            if ($InfoLevel.Networking -gt 0) {
                Write-PScriboMessage "Collecting VMM Networking information."
                Section -Style Heading1 'Networking' {
                    Paragraph 'The following section contains as built for Logical Networks, Logical Switches, Port Profiles and VM Networks'
                    Get-AbrVmmLogicalNetwork
                    Get-AbrVmmVmNetwork
                    Get-AbrVmmLogicalSwitch
                    Get-AbrVmmUplinkPortProfile
                    Get-AbrVmmNetworkAdapterPortProfile
                    Get-AbrVmmPortClassification
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}