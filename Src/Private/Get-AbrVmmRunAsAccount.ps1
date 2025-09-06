function Get-AbrVmmRunAsAccount {
    <#
    .SYNOPSIS
        Used by As Built Report to retrieve Microsoft VMM Server RunAs Account information
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
                if ($VmmRunAsAccounts = Get-SCRunAsAccount -VMMServer $ConnectVmmServer | Sort-Object -Property Name) {
                    Write-PScriboMessage "Collecting VMM Server RunAs Account information."
                    Section -Style Heading2 'RunAs Accounts' {
                        Paragraph "The following table summarises the configuration of the vmm server runas account."
                        BlankLine
                        $VmmRunAsAccountsInfo = @()
                        foreach ($VmmRunAsAccount in $VmmRunAsAccounts) {
                            $InObj = [Ordered]@{
                                'Name' = $VmmRunAsAccount.Name
                                'User Name' = $VmmRunAsAccount.UserName
                                'Domain' = $VmmRunAsAccount.Domain
                                'Enabled' = $VmmRunAsAccount.Enabled
                                'Is BuiltIn' = $VmmRunAsAccount.IsBuiltIn
                                'User Role' = $VmmRunAsAccount.UserRole
                            }

                            $VmmRunAsAccountsInfo += [pscustomobject](ConvertTo-HashToYN $InObj)
                        }

                        $TableParams = @{
                            Name = "RunAs Accounts - $($VMM.FQDN)"
                            List = $false
                            ColumnWidths = 20, 20, 18, 12, 12, 18
                        }
                        if ($Report.ShowTableCaptions) {
                            $TableParams['Caption'] = "- $($TableParams.Name)"
                        }
                        $VmmRunAsAccountsInfo | Table @TableParams
                    }
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $($_.Exception.Message)
        }
    }

    end {}
}