function Get-AbrVmmRequiredModule {
    <#
    .SYNOPSIS
        Function to check if the required version of VirtualMachineManager is installed
    .DESCRIPTION
        Documents the configuration of Veeam VMM in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.1.1
        Author:         Jonathan Colon
        Twitter:        @jcolonfzenpr
        Github:         rebelinux
        Credits:        Iain Brighton (@iainbrighton) - PScribo module

    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.Veeam.VMM
    #>
    [CmdletBinding()]

    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Version
    )
    process {
        #region: Start Load VirtualMachineManager Module
        # Loading Module
        # Make sure PSModulePath includes SCVMM Console
        #Code taken from @vMarkus_K
        if (Test-Path "C:\Program Files\Microsoft System Center\Virtual Machine Manager\bin\psModules" ) {
            $MyModulePath = "C:\Program Files\Microsoft System Center\Virtual Machine Manager\bin\psModules"
            $env:PSModulePath = $env:PSModulePath + "$([System.IO.Path]::PathSeparator)$MyModulePath"
        } elseif (Test-Path "D:\Program Files\Microsoft System Center\Virtual Machine Manager\bin\psModules" ) {
            $MyModulePath = "D:\Program Files\Microsoft System Center\Virtual Machine Manager\bin\psModules"
            $env:PSModulePath = $env:PSModulePath + "$([System.IO.Path]::PathSeparator)$MyModulePath"
        } elseif (Test-Path "E:\Program Files\Microsoft System Center\Virtual Machine Manager\bin\psModules" ) {
            $MyModulePath = "E:\Program Files\Microsoft System Center\Virtual Machine Manager\bin\psModules"
            $env:PSModulePath = $env:PSModulePath + "$([System.IO.Path]::PathSeparator)$MyModulePath"
        }
        if ($Modules = Get-Module -ListAvailable -Name VirtualMachineManager) {
            try {
                Write-PScriboMessage "Trying to import Microsoft SCVMM modules."
                $Modules | Import-Module -WarningAction SilentlyContinue
            } catch {
                Write-PScriboMessage -IsWarning "Failed to load Microsoft SCVMM modules"
            }
        }
        Write-PScriboMessage "Identifying Microsoft SCVMM module version."
        if ($Module = Get-Module -ListAvailable -Name virtualmachinemanager) {
            try {
                $script:SCVmmVersion = $Module.Version.ToString()
                Write-PScriboMessage "Using Microsoft SCVMM powershell module version $($SCVmmVersion)."
            } catch {
                Write-PScriboMessage -IsWarning "Failed to get Version from Module"
            }
        }
        # Check if the required version of VMware PowerCLI is installed
        $RequiredModule = Get-Module -ListAvailable -Name $Name
        $ModuleVersion = "{0}.{1}" -f $RequiredModule.Version.Major, $RequiredModule.Version.Minor
        if ($ModuleVersion -eq ".") {
            throw "$Name $Version or higher is required to run the Microsoft SCVMM As Built Report. Install the Microsoft SCVMM console that provide the required modules."
        }

        if ($ModuleVersion -lt $Version) {
            throw "$Name $Version or higher is required to run the Microsoft SCVMM As Built Report. Update the Microsoft SCVMM console that provide the required modules."
        }
    }
    end {}
}