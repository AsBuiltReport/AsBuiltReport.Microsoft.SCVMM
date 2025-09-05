function Get-AbrVmmLibraryTemplate {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Library and Templates information
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
        Write-PScriboMessage "Networking InfoLevel set at $($InfoLevel.LibraryTemplates)."
    }

    process {
        try {
            if ($InfoLevel.LibraryTemplates -gt 0) {
                Write-PScriboMessage "Collecting VMM Library and Template information."
                Section -Style Heading1 'Library and Templates' {
                    Paragraph 'The following section details the library and vm templates configured'
                    Get-AbrVmmLibraryServer
                    Get-AbrVmmLibraryShare
                    Get-AbrVmmVMTemplate
                    Get-AbrVmmGuestOSProfile
                    Get-AbrVmmHardwareProfile
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}