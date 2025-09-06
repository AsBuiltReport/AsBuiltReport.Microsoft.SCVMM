function Get-AbrVmmUserRole {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft VMM Server User Roles information
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
                if ($VmmUserRoles = Get-SCUserRole -VMMServer $ConnectVmmServer | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Server User Roles information."
                    Section -Style Heading2 'User Roles' {
                        $VmmUserRolesInfo = @()
                        foreach ($VmmUserRole in $VmmUserRoles) {
                            $InObj = [Ordered]@{
                                'Name' = $VmmUserRole.Name
                                'User Name' = $VmmUserRole.UserName
                                'Role Profile Type' = switch ($VmmUserRole.UserRoleProfile) {
                                    'Administrator' { 'Administrator' }
                                    'DelegatedAdmin' { 'Fabric Administrator' }
                                    'VMAdmin' { 'Virtual Machine Admin' }
                                    'SelfServiceUser' { 'Application Administrator' }
                                    'TenantAdmin' { 'Tenant Administrator' }
                                    'ReadOnlyAdmin' { 'ReadOnly Administrator' }
                                    Default { $VmmUserRole.UserRoleProfile }
                                }
                                'RunAs Account' = $VmmUserRole.RunAsAccount
                                'Library Server' = $VmmUserRole.LibraryServer
                                'Cloud' = $VmmUserRole.Cloud
                                'Parent User Role' = $VmmUserRole.ParentUserRole
                                'Members' = $VmmUserRole.Members.Name -join ", "
                                'Description' = $VmmUserRole.Description
                            }

                            $VmmUserRolesInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        if ($InfoLevel.Infrastructure -ge 2) {
                            Paragraph "The following sections detail the configuration of the vmm server user roles."
                            foreach ($VmmUserRole in $VmmUserRolesInfo) {
                                Section -Style NOTOCHeading4 -ExcludeFromTOC "$($VmmUserRole.'Name')" {
                                    $TableParams = @{
                                        Name = "User Roles - $($VmmUserRole.'Name')"
                                        List = $true
                                        ColumnWidths = 40, 60
                                    }
                                    if ($Report.ShowTableCaptions) {
                                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                    }
                                    $VmmUserRole | Table @TableParams
                                }
                            }
                        } else {
                            Paragraph "The following table summarises the configuration of the vmm server user roles."
                            BlankLine
                            $TableParams = @{
                                Name = "User Roles - $($VMM.FQDN)"
                                List = $false
                                Columns = 'Name', 'Description', 'Role Profile Type', 'Parent User Role'
                                ColumnWidths = 30, 30, 20, 20
                            }
                            if ($Report.ShowTableCaptions) {
                                $TableParams['Caption'] = "- $($TableParams.Name)"
                            }
                            $VmmUserRolesInfo | Table @TableParams
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