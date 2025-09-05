function Get-AbrVmmHost {
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
            if ($InfoLevel.Hosts -gt 0) {
                Write-PScriboMessage "Collecting VMM Host information."
                if ($ScVmmHosts = Get-SCVMHost) {
                    Section -Style Heading1 'Hosts' {
                        Paragraph "The following table summarises the configuration of the hosts."
                        BlankLine
                        Get-AbrVmmHostSummary
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}