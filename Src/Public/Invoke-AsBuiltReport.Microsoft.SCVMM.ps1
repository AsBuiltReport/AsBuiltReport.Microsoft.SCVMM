function Invoke-AsBuiltReport.Microsoft.SCVMM {
    <#
    .SYNOPSIS
        PowerShell script to document the configuration of Microsoft SCVMM in Word/HTML/Text formats
    .DESCRIPTION
        Documents the configuration of Microsoft SCVMM in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.1.1
        Author:         AsBuiltReport Organization
        Twitter:        @AsBuiltReport
        Github:         https://github.com/AsBuiltReport
        Credits:        Iain Brighton (@iainbrighton) - PScribo module

    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM
    #>

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", "", Scope = "Function")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingUserNameAndPassWordParams", "", Scope = "Function")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "", Scope = "Function")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Scope = "Function")]

    # Do not remove or add to these parameters
    param (
        [String[]] $Target,
        [PSCredential] $Credential
    )

    #Requires -Version 5.1
    #Requires -PSEdition Desktop
    #Requires -RunAsAdministrator

    if ($psISE) {
        Write-Error -Message "You cannot run this script inside the PowerShell ISE. Please execute it from the PowerShell Command Window."
        break
    }

    Get-AbrVmmRequiredModule -Name 'VirtualMachineManager' -Version '1.0'

    Write-Host "- Please refer to the AsBuiltReport.Microsoft.SCVMM github website for more detailed information about this project."
    Write-Host "- Do not forget to update your report configuration file after each new version release."
    Write-Host "- Documentation: https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM"
    Write-Host "- Issues or bug reporting: https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM/issues"
    Write-Host "- This project is community maintained and has no sponsorship from Microsoft, its employees or any of its affiliates."


    # Check the version of the dependency modules
    $ModuleArray = @('AsBuiltReport.Microsoft.SCVMM', 'Diagrammer.Core')

    foreach ($Module in $ModuleArray) {
        try {
            $InstalledVersion = Get-Module -ListAvailable -Name $Module -ErrorAction SilentlyContinue | Sort-Object -Property Version -Descending | Select-Object -First 1 -ExpandProperty Version

            if ($InstalledVersion) {
                Write-Host "- $Module module v$($InstalledVersion.ToString()) is currently installed."
                $LatestVersion = Find-Module -Name $Module -Repository PSGallery -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Version
                if ($InstalledVersion -lt $LatestVersion) {
                    Write-Host "  - $Module module v$($LatestVersion.ToString()) is available." -ForegroundColor Red
                    Write-Host "  - Run 'Update-Module -Name $Module -Force' to install the latest version." -ForegroundColor Red
                }
            }
        } catch {
            Write-PScriboMessage -IsWarning $_.Exception.Message
        }
    }

    # Import Report Configuration
    $Report = $ReportConfig.Report
    $InfoLevel = $ReportConfig.InfoLevel
    $Options = $ReportConfig.Options

    # Used to set values to TitleCase where required
    $TextInfo = (Get-Culture).TextInfo

    #region foreach loop
    foreach ($Server in $Target) {
        # Establish initial connection to VMM server
        Get-AbrVmmServerConnection

        # Variable translating Icon to Image Path ($IconPath)
        $script:Images = @{
            "AsBuiltReport_LOGO" = "AsBuiltReport_Logo.png"
            "AsBuiltReport_Signature" = "AsBuiltReport_Signature.png"
            "Abr_LOGO_Footer" = "AsBuiltReport_Signature.png"
            "Microsoft_Logo" = "Microsoft_Logo.png"
            "Server" = "server.png"
            "DB_Server" = "DB_Server.png"
        }
        $script:ColumnSize = $Options.DiagramColumnSize

        Write-Verbose "`VMM Server [$($VMM.name)] connection status is [$($VMM.IsConnected)]"
        Section -Style Heading1 $($VMM.FQDN) {
            Paragraph "The following section details the configuration of SCVMM server $($VMM.FQDN)."

            $VMMDiagram = Get-AbrVmmInfrastructureDiagram
            if ($VMMDiagram) {
                Export-AbrDiagram -DiagramObject $VMMDiagram -MainDiagramLabel "Infrastructure Diagram" -FileName "AsBuiltReport.Microsoft.SCVMM.Infrastructure"
            } else {
                Write-PScriboMessage -IsWarning "Unable to generate the Infrastructure Diagram."
            }

            Get-AbrVmmInfrastructure
            Get-AbrVmmNetworking
            Get-AbrVmmLibraryTemplate
            Get-AbrVmmCluster
            Get-AbrVmmHostNHostGroup
        }
    }
    #endregion foreach loop
}
