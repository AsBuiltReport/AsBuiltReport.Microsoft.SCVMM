function Get-AbrVmmLibraryShare {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Library Shares information
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
        Write-PScriboMessage "LibraryTemplates InfoLevel set at $($InfoLevel.LibraryTemplates)."
    }

    process {
        try {
            if ($InfoLevel.LibraryTemplates -gt 0) {
                if ($LibraryShares = Get-SCLibraryShare -VMMServer $ConnectVmmServer | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Library Shares information."
                    Section -Style Heading3 'Library Shares' {
                        $VmmLibrarySharesInfo = @()
                        foreach ($LibraryShare in $LibraryShares) {
                            $InObj = [Ordered]@{
                                'Name' = $LibraryShare.Name
                                'Library Server' = $LibraryShare.LibraryServer
                                'Path' = $LibraryShare.Path
                                'Use Alternate DataStream?' = $LibraryShare.UseAlternateDataStream
                                'Description' = $LibraryShare.Description
                            }

                            $VmmLibrarySharesInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.LibraryTemplates -ge 2) {
                            Paragraph "The following sections detail the configuration of the library shares."
                            foreach ($LibraryShare in $VmmLibrarySharesInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($LibraryShare.Name)" {
                                    $TableParams = @{
                                        Name = "Library Shares - $($LibraryShare.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $LibraryShare | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the library shares."
                            BlankLine
                            $TableParams = @{
                                Name = "Library Shares - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'Library Server', 'Path', 'Use Alternate DataStream?'
                                ColumnWidths = 25, 30, 30, 15
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmLibrarySharesInfo | Table @TableParams
                        }
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}