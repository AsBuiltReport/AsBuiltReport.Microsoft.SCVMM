function Get-AbrVmmGuestOSProfile {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft SCVMM Guest OS Profile information
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
                if ($GuestOSProfiles = Get-SCGuestOSProfile -VMMServer $ConnectVmmServer | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Guest OS Profiles information."
                    Section -Style Heading3 'Guest OS Profiles' {
                        $VmmGuestOSProfilesInfo = @()
                        foreach ($GuestOSProfile in $GuestOSProfiles) {
                            $InObj = [Ordered]@{
                                'Name' = $GuestOSProfile.Name
                                'Full Name' = $GuestOSProfile.FullName
                                'Operating System' = $GuestOSProfile.OperatingSystem
                                'Operating System Type' = $GuestOSProfile.OSType
                                'Computer Name' = $GuestOSProfile.ComputerName
                                'Product Key' = $GuestOSProfile.ProductKey
                                'Join Workgroup' = $GuestOSProfile.JoinWorkgroup
                                'Join Domain RunAs Account' = $GuestOSProfile.JoinDomainRunAsAccount
                                'Join Domain' = $GuestOSProfile.JoinDomain
                                'Admin Password RunAsAccount' = $GuestOSProfile.AdminPasswordRunAsAccount
                                'Org Name' = $GuestOSProfile.OrgName
                                'Domain Admin' = $GuestOSProfile.DomainAdmin
                                'DNS Domain Name' = $GuestOSProfile.DNSDomainName
                                'Sysprep Script' = $GuestOSProfile.SysprepScript
                                'Domain Join Organizational Unit' = $GuestOSProfile.DomainJoinOrganizationalUnit
                                'Server Features' = $GuestOSProfile.ServerFeatures.DisplayName
                                'Description' = $GuestOSProfile.Description
                            }

                            $VmmGuestOSProfilesInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.LibraryTemplates -ge 2) {
                            Paragraph "The following sections detail the configuration of the guest os profiles."
                            foreach ($GuestOSProfile in $VmmGuestOSProfilesInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($GuestOSProfile.Name)" {
                                    $TableParams = @{
                                        Name = "Guest OS Profiles - $($GuestOSProfile.Name)"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $GuestOSProfile | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the guest os profiles."
                            BlankLine
                            $TableParams = @{
                                Name = "Guest OS Profiles - $($Vmm.FQDN)"
                                List = $false
                                Columns = 'Name', 'Join Domain', 'Operating System Type', 'Server Features'
                                ColumnWidths = 25, 20, 20, 35
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmGuestOSProfilesInfo | Table @TableParams
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