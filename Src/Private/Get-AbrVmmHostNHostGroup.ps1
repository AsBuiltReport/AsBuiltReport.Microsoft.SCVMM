function Get-AbrVmmHostNHostGroup {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Hosts information
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
        Write-PScriboMessage "Networking InfoLevel set at $($InfoLevel.Hosts)."
    }

    process {
        try {
            if ($InfoLevel.Hosts -gt 0 -or $InfoLevel.HostGroups -gt 0) {
                Write-PScriboMessage "Collecting VMM Host and Host Group information."
                Section -Style Heading1 'Host Groups & Hosts' {
                    Paragraph "The following table summarises the configuration of host groups and hosts."
                    BlankLine
                    Get-AbrVmmHostGroup
                    Get-AbrVmmHost
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}