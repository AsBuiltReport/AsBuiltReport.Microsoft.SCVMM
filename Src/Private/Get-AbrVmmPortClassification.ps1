function Get-AbrVmmPortClassification {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Port Classification information
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
                if ($PortClassifications = Get-SCPortClassification | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Port Classification information."
                    Section -Style Heading3 'Port Classifications' {
                        $VmmPortClassificationInfo = @()
                        foreach ($PortClassification in $PortClassifications) {
                            $InObj = [Ordered]@{
                                'Name' = $PortClassification.Name
                                'Description' = $PortClassification.Description
                            }

                            $VmmPortClassificationInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        Paragraph "The following table summarises the configuration of the port classification."
                        BlankLine
                        $TableParams = @{
                            Name = "Port Classification - $($Vmm.FQDN)"
                            List = $false
                            Columns = 'Name', 'Description'
                            ColumnWidths = 40, 60
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $VmmPortClassificationInfo | Sort-Object -Property Name | Table @TableParams
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}