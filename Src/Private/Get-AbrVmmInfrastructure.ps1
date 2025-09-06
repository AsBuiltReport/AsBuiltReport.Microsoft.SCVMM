function Get-AbrVmmInfrastructure {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft VMM Server information
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
        Write-PScriboMessage "Infrastructure InfoLevel set at $($InfoLevel.Infrastructure)."
    }

    process {
        try {
            if ($InfoLevel.Infrastructure -gt 0) {
                if ($Vmm) {
                    Write-PScriboMessage "Collecting VMM Server Settings information."
                    Section -Style Heading2 'Infrastructure' {
                        Get-AbrVmmServer
                        Get-AbrVmmDBSetting
                        Get-AbrVmmAutoNetSetting
                        Get-AbrVmmUpdateServer
                        Get-AbrVmmLibraryServer
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}