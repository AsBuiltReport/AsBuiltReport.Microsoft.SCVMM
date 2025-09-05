function Get-AbrVmmVMTemplate {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM VM Templates information
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
                if ($VMTemplates = Get-SCVMTemplate | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM VM Templates information."
                    Section -Style Heading3 'VM Templates' {
                        $VmmVMTemplatesInfo = @()
                        foreach ($VMTemplate in $VMTemplates) {
                            $InObj = [Ordered]@{
                                'Name' = $VMTemplate.Name
                                'Operating System' = $VMTemplate.OperatingSystem
                                'CPU Count' = $VMTemplate.CPUCount
                                'Product Key' = $VMTemplate.ProductKey
                                'Memory (MB)' = $VMTemplate.Memory
                                'JoinWorkgroup' = $VMTemplate.JoinWorkgroup
                                'OrgName' = $VMTemplate.OrgName
                                'DomainAdmin' = $VMTemplate.DomainAdmin
                                'ComputerName' = $VMTemplate.ComputerName
                                'FullName' = $VMTemplate.FullName
                                'DNSDomainName' = $VMTemplate.DNSDomainName
                                'SysprepScript' = $VMTemplate.SysprepScript
                                'DynamicMemoryEnabled' = $VMTemplate.DynamicMemoryEnabled
                                'VirtualVideoAdapterEnabled' = $VMTemplate.VirtualVideoAdapterEnabled
                                'MonitorMaximumCount' = $VMTemplate.MonitorMaximumCount
                                'MonitorResolutionMaximum' = $VMTemplate.MonitorResolutionMaximum
                                'UseHardwareAssistedVirtualization' = $VMTemplate.UseHardwareAssistedVirtualization
                                'Tags' = ($VMTemplate.Tags -join ', ')
                                'CapabilityProfile' = $VMTemplate.CapabilityProfile
                                'VirtualizationPlatform' = $VMTemplate.VirtualizationPlatform
                                'DomainJoinOrganizationalUnit' = $VMTemplate.DomainJoinOrganizationalUnit
                                'Generation' = $VMTemplate.Generation
                                'Description' = $VMTemplate.Description
                            }

                            $VmmVMTemplatesInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.LibraryTemplates -ge 2) {
                            Paragraph "The following sections detail the configuration of the vm templates."
                            foreach ($VMTemplate in $VmmVMTemplatesInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($VMTemplate.Name)" {
                                    $TableParams = @{
                                        Name = "VM Templates - $($VMTemplate.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VMTemplate | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the vm templates."
                            BlankLine
                            $TableParams = @{
                                Name = "VM Templates - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'Operating System', 'CPU Count', 'Memory (MB)', 'Generation'
                                ColumnWidths = 27, 28, 15, 15, 15
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmVMTemplatesInfo | Table @TableParams
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