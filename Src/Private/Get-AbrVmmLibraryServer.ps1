function Get-AbrVmmLibraryServer {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Library Servers information
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
                if ($VMLibraries = Get-SCLibraryServer | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Library Servers information."
                    Section -Style Heading3 'Library Servers' {
                        $VmmLibraryServersInfo = @()
                        foreach ($VMLibrary in $VMLibraries) {
                            $InObj = [Ordered]@{
                                'Name' = $VMLibrary.Name
                                'FQDN' = $VMLibrary.FQDN
                                'Status' = $VMLibrary.Status
                                'Domain Name' = $VMLibrary.DomainName
                                'Library Group' = $VMLibrary.LibraryGroup
                                'Is Cluster Node?' = $VMLibrary.IsClusterNode
                                'Is Virtual Cluster Name?' = $VMLibrary.IsVirtualClusterName
                                'VM Networks' = ($VMLibrary.VMNetworks | ForEach-Object { $_.Name }) -join ', '
                                'Is Unencrypted File Transfer Enabled?' = $VMLibrary.IsUnencryptedFileTransferEnabled
                                'Library Server Management Credential' = $VMLibrary.LibraryServerManagementCredential.Name
                                'Description' = $VMLibrary.Description
                            }

                            $VmmLibraryServersInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.LibraryTemplates -ge 2) {
                            Paragraph "The following sections detail the configuration of the library servers."
                            foreach ($LibraryServer in $VmmLibraryServersInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($LibraryServer.Name)" {
                                    $TableParams = @{
                                        Name = "Library Servers - $($LibraryServer.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $LibraryServer | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the library servers."
                            BlankLine
                            $TableParams = @{
                                Name = "Library Servers - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'Description', 'Status'
                                ColumnWidths = 42, 42, 16
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmLibraryServersInfo | Table @TableParams
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