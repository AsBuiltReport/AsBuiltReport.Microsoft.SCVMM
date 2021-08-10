function Invoke-AsBuiltReport.Microsoft.SCVMM {
    <#
    .SYNOPSIS
        PowerShell script to document the configuration of Microsoft SCVMM in Word/HTML/Text formats
    .DESCRIPTION
        Documents the configuration of Microsoft SCVMM in Word/HTML/Text formats using PScribo.
    .NOTES
        Version:        0.1.0
        Author:         Andrew Ramsay
        Twitter:
        Github:
        Credits:        Iain Brighton (@iainbrighton) - PScribo module

    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM
    #>

	# Do not remove or add to these parameters
    param (
        [String[]] $Target,
        [PSCredential] $Credential
    )

    # Import Report Configuration
    $Report = $ReportConfig.Report
    $InfoLevel = $ReportConfig.InfoLevel
    $Options = $ReportConfig.Options

    # Used to set values to TitleCase where required
    $TextInfo = (Get-Culture).TextInfo

	# Update/rename the $VmmServer variable and build out your code within the ForEach loop. The ForEach loop enables AsBuiltReport to generate an as built configuration against multiple defined targets.

    #region foreach loop
    foreach ($Server in $Target) {
        $ConnectVmmServer = Get-SCVMMServer -ComputerName $Server -Credential $Credential
		Write-Verbose "`VMM Server [$($ConnectVmmServer.name)] connection status is [$($ConnectVmmServer.IsConnected)]"
        Section -Style Heading1 'Virtual Machine Manager Server' {
            $VMM = $ConnectVmmServer
            $vmmCim = New-CimSession -ComputerName ($VMM.FQDN) -Credential $Credential
            $VMMFQDN = $VMM.FQDN

            Paragraph "The following section details the configuration of SCVMM server $VMMFQDN."

            Section -Style Heading2 $VMMFQDN  {
                $VMMServerSettingsReport = [PSCustomObject]@{
                    'Server FQDN' = $VMMFQDN
                    'IP Address' = (Get-NetIPAddress -CimSession $vmmCim -AddressFamily IPv4 | Where-Object {$_.IPAddress -notlike "127.0.0.1"}).IPAddress
                    'Product Version' = $VMM.ProductVersion
                    'Server Port' = $VMM.Port
                    'VM Connect Port' = $VMM.VMConnectDefaultPort
                    'VMM Service Account' = $VMM.VMMServiceAccount
                    'VMM High Availability' = $VMM.IsHighlyAvailable
                }
                $TableParams = @{
                    Name = 'VMM Server Settings'
                    List = $true
                    ColumnWidths = 50,50
                }
                $VMMServerSettingsReport | Table $TableParams
            }
            Section -Style Heading3 'VMM Database Settings' {
                $VMMDBSettingsReport = [PSCustomObject]@{
                    'DB Server Name' = $VMM.DatabaseServerName
                    'DB Instance Name' = $VMM.DatabaseInstanceName
                    'DB Name' = $VMM.DatabaseName
                    'DB Version' = $VMM.DatabaseVersion
                }
                $VMMDBSettingsReport | Table -Name 'VMM DB Settings' -List -ColumnWidths 50,50
            }
            Section -Style Heading3 'VMM AutoNetwork Settings' {
                $VmmAutoNetworkSettingsReport = [PSCustomObject]@{
                    'Logical Network Creation Enabled' = $VMM.AutomaticLogicalNetworkCreationEnabled
                    'Virtual Network Creation Enabled' = $VMM.AutomaticVirtualNetworkCreationEnabled
                    'Logical Network Match' = $VMM.LogicalNetworkMatchOption
                    'Backup Network Match' = $VMM.BackupLogicalNetworkMatchOption
                }
                $VMMAutoNetworkSettingsReport | Table -Name 'VMM Server Settings' -List -ColumnWidths 50,50
            }
        }
        Section -Style Heading1 'VMM Networking' {
            Paragraph 'The following section contains as built for Logical Networks, Logical Switches, Port Profiles and VM Networks'
            Section -Style Heading2 'Logical Networks' {
                $LogicalNetworks = Get-SCLogicalNetworkDefinition
                Paragraph 'The summary for logical networks is as follws'
                $LogicalNetworks | Select-Object Name,IsolationType | Table -Name 'Logical Network Summary'
                foreach($network in $LogicalNetworks){
                    Section -Style Heading3 $network.Name {
                        $network | Select-Object Name,IsolationType,@{L="HostGroups";E={$_.HostGroups -Join ","}}| Table -Name $network.Name
                    }
                    Section -Style Heading3 'Vlans and Subnets' {
                        $network.SubnetVLans | Select-Object VLanID,Subnet | Table -Name ($network.Name + "VLANs")
                    }
                }
            }
            Section -Style Heading2 'VM Networks' {
                Paragraph 'The following section details VM Networks'
                BlankLine
                $VMNetworks = Get-SCVMNetwork
                $VMNetworkReport = @()
                ForEach($VMNetwork in $VMNetworks) {
                    $TempVMNetworkReport = [PSCustomObject]@{
                        'Name' = $VMNetwork.Name
                        'Logical Network' = $VMNetwork.LogicalNetwork
                        'VLAN' = $VMNetwork.VMSubnet.SubnetVLans.VLanID
                        'Subnet' = $VMNetwork.VMSubnet.SubnetVLans.Subnet
                        'Isolation Type' = $VMNetwork.IsolationType
                    }
                    $VMNetworkReport += $TempVMNetworkReport
                }
                $VMNetworkReport | Table -Name 'VM Networks'
            }
            Section -Style Heading2 'Logical Switches'{
                Paragraph 'The following section contains as-built for Logical Switches'
                $LogicalSwitches = Get-SCLogicalSwitch
                if($LogicalSwitches){
                    $LogicalSwitchesReport = @()
                    ForEach($TempPssSessionwitch in $LogicalSwitches){
                        $TempLogicalSwitchesReport = [PSCustomObject]@{
                            'Name' = $TempPssSessionwitch.Name
                            'Uplink Mode' = $TempPssSessionwitch.UplinkMode
                            'Minimum Bandwidth Mode' = $TempPssSessionwitch.MinimumMandwidthMode
                        }
                        $LogicalSwitchesReport += $TempLogicalSwitchesReport
                    }
                    $LogicalSwitchesReport | Table -Name 'LogicalSwitches'
                }
            }
            Section -Style Heading2 'Uplink Port Profiles'{
                Paragraph 'The following section contains as-built for Uplink Port Profiles'
                $UplinkPortProfile = Get-SCNativeUplinkPortProfile
                if($UplinkPortProfle){
                    $UplinkReport = $UplinkPortProfile | Select-Object Name, `
                    @{L='Team Mode'; E={$_.LBFOTeamMode}}, `
                    @{L='Load Balance'; E={$_LBFOLoadBalancingAlgorithm}}
                    $UplinkReport | Table -Name 'Uplink Port Profiles'
                }
            }
            Section -Style Heading2 'Network Adapter Port Profiles' {
                Paragraph 'The following section details Virtual Network Adapter Port Profiles'
                $VirtualNetworkAdapterPortProfiles = Get-SCVirtualNetworkAdapterNativePortProfile
                if($VirtualNetworkAdapterPortProfiles){
                    foreach($adapter in $VirtualNetworkAdapterPortProfiles){
                        Section -Style Heading3 ($adapter.Name) {
                            Paragraph ($adapter.Description)
                            $AdapterReport = [PSCustomObject]@{
                                'Teaming' = $adapter.AllowTeaming
                                'Mac Address Spoofing' = $adapter.AllowMacAddressSpoofing
                                'Ieee Priority Tagging' = $adapter.AllowIeeePriorityTagging
                                'DHCP Guard' = $adapter.EnableDHCPGuard
                                'Guest IP Network Virtualization Updates' = $adapter.EnableGuestIPNetworkVirtualizationUpdates
                                'Router Guard' = $adapter.EnableRouterGuard
                                'Minimum Bandwidth Weight' = $adapter.MinimumBandwidthWeight
                                'Minimum Bandwidth Absolute In Mbps' = $adapter.MinimumBandwidthAbsoluteInMbps
                                'Maximum Bandwidth Absolute In Mbps' = $adapter.MaximumBandwidthAbsoluteInMbps
                                'Enable Vmq' = $adapter.EnableVmq
                                'Enable IPsec Offload' = $adapter.EnableIPsecOffload
                                'Enable Iov' = $adapter.EnableIov
                                'Enable Vrss' = $adapter.EnableVrss
                                'Enable Rdma' = $adapter.EnableRdma
                            }
                            $AdapterReport | Table -Name ($adapter.'Name') -List -ColumnWidths 50,50
                        }
                    }
                }
            }
            Section -Style Heading2 'Port Classifications' {
                Paragraph 'The following section details Port Classifications Configured'
                $PortClassifications = Get-SCPortClassification
                $PortClassifications | Select-Object Name, Description | Table -Name 'Port Classifications'
            }
        }
        Section -Style Heading1 'VMM Library and Templates'{
            Paragraph 'The following section details the Library and VM Templates configured'
            Section -Style Heading2 'VMM Library Servers'{
                Paragraph 'The following table is a summary of the VMM Library servers deployed'
                $VMLibrary = Get-SCLibraryServer
                $VMLibrary | Select-Object ComputerName,Description,Status | Table -Name 'VMM Library Servers'
            }
            Section -Style Heading2 'VMM Library Shares'{
                Paragraph 'The following table details the Library Shares Configured'
                $VMLibraryShares = Get-SCLibraryShare
                $VMLibraryShares | Select-Object Name,Description,LibraryServer | Table -Name 'VMM Library Shares'
            }
            Section -Style Heading2 'VMM Templates'{
                Paragraph 'The following table is a summary of VM Templates Deployed'
                $VMTemplates = Get-SCVMTemplate
                if($VMTemplates){
                    $VMTemplates | Select-Object Name,OperatingSystem,Description | Table -Name 'VMM Templates' -ColumnWidths 20,20,60
                }
            }
            Section -Style Heading2 'Guest OS Profiles'{
                Paragraph 'The following table is a summary of the Guest OS Profiles Deployed'
                $GuestOSProfiles = Get-SCGuestOsProfile
                $GuestOSProfiles | Select-Object Name,JoinDomain,OSTYpe | Table -Name 'Guest OS Profiles'
            }
            Section -Style Heading2 'Hardware Profiles'{
                Paragraph 'The following table is a summary of deployed Hardware Profiles'
                $HardwareProfiles = Get-SCHardwareProfile
                if($HardwareProfiles){
                    $HardwareProfiles | Select-Object Name,CPUCount,Memory,IsHighlyAvailable,SecureBootEnabled | Table -Name 'Hardware Profiles'
                }
            }
        }
        Section -Style Heading1 'Clusters' {
            Paragraph 'The following section details Hyper-V Clusters'
            $ScVmmClusters = Get-SCVMHostCluster
            #$Clusters = Get-Cluster -Name $TempPssSessioncVmmClusters.Name
            Paragraph 'The following table is the summary of the clusters managed by SCVMM'
            $ClusterSummaryReport = @()
            ForEach($ScVmmCluster in $ScVmmClusters){
                $TempReport = [PSCustomObject]@{
                    'Cluster Name' = $ScVmmCluster.Name
                    'Cluster IP' = $ScVmmCluster.IPAddresses -Join ","
                    'Host Group' = $ScVmmCluster.HostGroup
                    'Cluster Nodes' = $ScVmmCluster.Nodes -Join ","
                }
                $ClusterSummaryReport += $TempReport
            }
            $ClusterSummaryReport | Table -Name 'Cluster Summary' -ColumnWidths 20,10,30,40
            Section -Style Heading2 'Cluster Details' {
                ForEach($ScVmmCluster in $ScVmmClusters){
                    $Cluster = Get-Cluster -Name $ScVmmCluster.Name
                    Section -Style Heading3 $Cluster.Name {
                        #region Cluster Settings
                        Section -Style Heading4 'Cluster Settings' {
                            $ClusterSettingsTable = [PSCustomObject] @{
                                'Add Evict Delay' = $Cluster.AddEvictDelay
                                'Administrative Access Point' = $Cluster.AdministrativeAccessPoint
                                'Auto Assign Node Site' = $Cluster.AutoAssignNodeSite
                                'Auto Balancer Mode' = $Cluster.AutoBalancerMode
                                'Auto Balancer Level' = $Cluster.AutoBalancerLevel
                                'Backup In Progress' = $Cluster.BackupInProgress
                                'Block Cache Size' = $Cluster.BlockCacheSize
                                'Cluster Service Hang Timeout' = $Cluster.ClusSvcHangTimeout
                                'Cluster Service Regroup Stage Timeout' = $Cluster.ClusSvcRegroupStageTimeout
                                'Cluster Service Regroup Tick In Milliseconds' = $Cluster.ClusSvcRegroupTickInMilliseconds
                                'Cluster Enforced AntiAffinity' = $Cluster.ClusterEnforcedAntiAffinity
                                'Cluster Functional Level' = $Cluster.ClusterFunctionalLevel
                                'Cluster Upgrade Version' = $Cluster.ClusterUpgradeVersion
                                'Cluster Group Wait Delay' = $Cluster.ClusterGroupWaitDelay
                                'Cluster Log Level' = $Cluster.ClusterLogLevel
                                'Cluster Log Size' = $Cluster.ClusterLogSize
                                'Cross Site Delay' = $Cluster.CrossSiteDelay
                                'Cross Site Threshold' = $Cluster.CrossSiteThreshold
                                'Cross Subnet Delay' = $Cluster.CCrossSubnetDelay
                                'Cross Subnet Threshold' = $Cluster.CrossSubnetThreshold
                                'Csv Balancer' = $Cluster.CsvBalancer
                                'Database Read Write Mode' = $Cluster.DatabaseReadWriteMode
                                'Default Network Role' = $Cluster.DefaultNetworkRole
                                'Description' = $Cluster.Description
                                'Domain' = $Cluster.Domain
                                'Drain On Shutdown' = $Cluster.DrainOnShutdown
                                'Dump Policy' = $Cluster.DumpPolicy
                                'Dynamic Quorum' = $Cluster.DynamicQuorum
                                'Enable Shared Volumes' = $Cluster.EnableSharedVolumes
                                'Fix Quorum' = $Cluster.FixQuorum
                                'Group Dependency Timeout' = $Cluster.GroupDependencyTimeout
                                'Hang Recovery Action' = $Cluster.HangRecoveryAction
                                'Ignore Persistent State On Startup' = $Cluster.IgnorePersistentStateOnStartup
                                'Log Resource Controls' = $Cluster.LogResourceControls
                                'Lower Quorum Priority Node Id' = $Cluster.LowerQuorumPriorityNodeId
                                'Message Buffer Length' = $Cluster.MessageBufferLength
                                'Minimum Never Preempt Priority' = $Cluster.MinimumNeverPreemptPriority
                                'Minimum Preemptor Priority' = $Cluster.MinimumPreemptorPriority
                                'Name' = $Cluster.Name
                                'Net ft IPSec Enabled' = $Cluster.NetftIPSecEnabled
                                'Placement Options' = $Cluster.PlacementOptions
                                'Plumb All Cross Subnet Routes' = $Cluster.PlumbAllCrossSubnetRoutes
                                'Preferred Site' = $Cluster.PreferredSite
                                'Prevent Quorum' = $Cluster.PreventQuorum
                                'Quarantine Duration' = $Cluster.QuarantineDuration
                                'Quarantine Threshold' = $Cluster.QuarantineThreshold
                                'Quorum Arbitration Time Max' = $Cluster.QuorumArbitrationTimeMax
                                'Recent Events Reset Time' = $Cluster.RecentEventsResetTime
                                'Request Reply Timeout' = $Cluster.RequestReplyTimeout
                                'Resiliency Default Period' = $Cluster.ResiliencyDefaultPeriod
                                'Resiliency Level' = $Cluster.ResiliencyLevel
                                'Route History Length' = $Cluster.RouteHistoryLength
                                'Same Subnet Delay' = $Cluster.SameSubnetDelay
                                'Same Subnet Threshold' = $Cluster.SameSubnetThreshold
                                'Security Level' = $Cluster.SecurityLevel
                                'Shared Volume Compatible Filters' = $Cluster.SharedVolumeCompatibleFilters
                                'Shared Volume Incompatible Filters' = $Cluster.SharedVolumeIncompatibleFilters
                                'Shared Volume Security Descriptor' = $Cluster.SharedVolumeSecurityDescriptor
                                'Shared Volumes Root' = $Cluster.SharedVolumesRoot
                                'Shared Volume VssWriter Operation Timeout' = $Cluster.SharedVolumeVssWriterOperationTimeout
                                'Shutdown Timeout In Minutes' = $Cluster.ShutdownTimeoutInMinutes
                                'Use Client Access Networks For Shared Volumes' = $Cluster.UseClientAccessNetworksForSharedVolumes
                                'Witness Database Write Timeout' = $Cluster.WitnessDatabaseWriteTimeout
                                'Witness Dynamic Weight' = $Cluster.WitnessDynamicWeight
                                'Witness Restart Interval' = $Cluster.WitnessRestartInterval
                            }
                            $ClusterSettingsTable | Table -Name 'Cluster Settings' -List -ColumnWidths 50,50
                        }
                        #end region Cluster Settings
                        #Cluster Nodes
                        Section -Style Heading4 'Cluster Nodes' {
                            Paragraph 'The following Nodes are members of the cluster'
                            $ClusterNodes = $Cluster | Get-ClusterNode
                            $ClusterNodeReport = @()
                            foreach($ClusterNode in $ClusterNodes){
                                $Temp = [PSCustomObject] @{
                                    'Name' = $ClusterNode.Name
                                    'Status Information' = $ClusterNode.StatusInformation
                                    'Node Weight' = $ClusterNode.NodeWeight
                                    'Model' = $ClusterNode.Model
                                    'Manufacturer' = $ClusterNode.Manufacturer
                                    'Serial Number' = $ClusterNode.SerialNumber
                                }
                                $ClusterNodeReport += $Temp
                            }
                            $ClusterNodeReport | Table -Name 'Cluster Nodes'
                        }
                        #Cluster Quorum
                        Section -Style Heading4 'Cluster Quorum' {
                            Paragraph 'The following Cluster Quorum Settings are applied'
                            $ClusterQuorum = $Cluster | Get-ClusterQuorum
                            $QuorumReport = [PSCustomObject] @{
                                'Name' = $ClusterQuorum.QuorumResource.Name
                                'State' = $ClusterQuorum.QuorumResource.State
                                'Owner Node' = $ClusterQuorum.QuorumResource.OwnerNode
                                'ResourceType' = $ClusterQuorum.QuorumResource.ResourceType
                            }
                            $QuorumReport | Table -Name 'Cluster Quorum Settings' -List -ColumnWidths 50,50
                        }
                        #Cluster Networks
                        Section -Style Heading4 'Cluster Networks'{
                            Paragraph 'The following Cluster Networks are configured'
                            $ClusterNetworks = $Cluster | Get-ClusterNetwork
                            $ClusterNetworkReport = @()
                            foreach($ClusterNetwork in $ClusterNetworks){
                                $TempClusterNetwork = [PSCustomObject]@{
                                    'Name' = $ClusterNetwork.Name
                                    'Description' = $ClusterNetwork.Description
                                    'Role' = $ClusterNetwork.Role
                                    'Network Address' = $ClusterNetwork.Address
                                    'State' = $ClusterNetwork.State
                                }
                                $ClusterNetworkReport += $TempClusterNetwork
                            }
                            $ClusterNetworkReport | Table -Name 'Cluster Networks'
                            #Cluster Network Interfaces
                            Section -Style Heading5 'Cluster Network Interfaces'{
                                Paragraph 'The following table details the network interfaces nodes use in the cluster'
                                $ClusterNetworkInterfaces = $Cluster | Get-ClusterNetworkInterface
                                $ClusterInterfaceReport = @()
                                foreach($Interface in $ClusterNetworkInterfaces){
                                    $TempInterfaceReport = [PSCustomObject]@{
                                        'Node' = $Interface.Node
                                        'Address' = $Interface.Address
                                        'Network' = $Interface.Network
                                        'Interface Name' = $Interface.Name
                                    }
                                    $ClusterInterfaceReport += $TempInterfaceReport
                                }
                                $ClusterInterfaceReport | Table -Name 'Cluster Interfaces'
                            }
                        }
                        #Cluster Storage
                        Section -Style Heading4 'Cluster Shared Volumes'{
                            Paragraph 'The following Cluster Shared Volumes are Configure'
                            $ClusterVolumes = $Cluster | Get-ClusterSharedVolume
                            if($ClusterVolumes){
                                $ClusterVolumeReport = @()
                                foreach($ClusterVolume in $ClusterVolumes){
                                    $TempClusterVolume = [PSCustomObject] @{
                                        'Name' = $ClusterVolume.Name
                                        'State' = $ClusterVolume.State
                                        'File System Type' = $ClusterVolume.SharedVolumeInfo.Partition.FileSystem
                                        'Volume Capacity(GB)' = [Math]::Round(($ClusterVolume.SharedVolumeInfo.Partition.Size)/1gb)
                                        'Free Fapacity(GB)' = [Math]::Round(($ClusterVolume.SharedVolumeInfo.Partition.FreeSpace)/1gb)
                                    }
                                    $ClusterVolumeReport += $TempClusterVolume
                                }
                                $ClusterVolumeReport | Table -Name 'Cluster Volumes'
                            }
                        }
                        #Cluster Hyper-V Replica Broker
                    }
                }
            }
        }
        Section -Style Heading1 'Hyper-V Hosts'{
            Paragraph 'The following table details the Hyper-V hosts'
            $VMHosts = Get-SCVMHost -VMMServer $ConnectVmmServer | Sort-Object Name
            $VMHostSummary = $VMHosts | Select-Object ComputerName,OperatingSystem,VMHostGroup
            $VMHostSummary | Table -Name 'VM Host Summary'
            #Host Summary
            ForEach($VMHost in $VMHosts){
                #Create Remote Sessions
                $TempPssSession = New-PSSession $VMHost.Name -Credential $Credential
                $TempCimSession = New-CimSession $VMHost.Name -Credential $Credential
                #Get Server Data using WinRM
                $HostInfo = Invoke-Command -Session $TempPssSession {Get-ComputerInfo}
                $HostCPU = Get-CimInstance -CimSession $TempCimSession -ClassName Win32_Processor
                $HostComputer = Get-CimInstance -CimSession $TempCimSession -ClassName Win32_ComputerSystem
                $HostBIOS = Get-CimInstance -CimSession $TempCimSession -ClassName Win32_Bios
                $HostLicense = Get-CimInstance -CimSession $TempCimSession -query 'Select * from SoftwareLicensingProduct'| Where-Object {$_.LicenseStatus -eq 1}
                $HotFixes = Get-CimInstance -CimSession $TempCimSession -ClassName Win32_QuickFixEngineering
                Section -Style Heading2 ($VMHost.Name){
                    #Host Hardware
                    Section -Style Heading3 'Host Hardware Settings'{
                        Paragraph 'The following section details hardware settings for the host'
                        $HostHardware = [PSCustomObject] @{
                            'Manufacturer' = $HostComputer.Manufacturer
                            'Model' = $HostComputer.Model
                            'Product ID' = $HostComputer.SystemSKUNumber
                            'Serial Number' = $HostBIOS.SerialNumber
                            'BIOS Version' = $HostBIOS.Version
                            'Processor Manufacturer' = $HostCPU[0].Manufacturer
                            'Processor Model' = $HostCPU[0].Name
                            'Number of Processors' = $HostCPU.Length
                            'Number of CPU Cores' = $HostCPU[0].NumberOfCores
                            'Number of Logical Cores' = $HostCPU[0].NumberOfLogicalProcessors
                            'Physical Memory (GB)' = [Math]::Round($HostComputer.TotalPhysicalMemory/1Gb)
                        }
                        $HostHardware | Table -Name 'Host Hardware Specifications' -List -ColumnWidths 50,50
                    }
                    #Host OS
                    Section -Style Heading3 'Host OS' {
                        Paragraph 'The following settings details host OS Settings'
                        Section -Style Heading4 'OS Configuration'{
                            Paragraph 'The following section details hos OS configuration'
                            $HostOSReport = [PSCustomObject] @{
                                'Windows Product Name' = $HostInfo.WindowsProductName
                                'Windows Version' = $HostInfo.WindowsCurrentVersion
                                'Windows Build Number' = $HostInfo.OsVersion
                                'Windows Install Type' = $HostInfo.WindowsInstallationType
                                'AD Domain' = $HostInfo.CsDomain
                                'Windows Installation Date' = $HostInfo.OsInstallDate
                                'Time Zone' = $HostInfo.TimeZone
                                'License Type' = $HostLicense.ProductKeyChannel
                                'Partial Product Key' = $HostLicense.PartialProductKey
                            }
                            $HostOSReport | Table -Name 'Host OS Settings' -List -ColumnWidths 50,50
                        }
                        Section -Style Heading4 'Host Hotfixes'{
                            Paragraph 'The following table details the OS Hotfixes installed'
                            $HotFixReport = @()
                            Foreach($HotFix in $HotFixes){
                                $TempHotFix = [PSCustomObject] @{
                                    'Hotfix ID' = $HotFix.HotFixID
                                    'Description' = $HotFix.Description
                                    'Installation Date' = $HotFix.InstalledOn
                                }
                                $HotFixReport += $TempHotFix
                            }
                            $HotFixReport | Table -Name 'HostFixes Installed' -ColumnWidths 10,70,20
                        }
                        Section -Style Heading4 'Host Drivers'{
                            Paragraph 'The following section details host drivers'
                            Invoke-Command -Session $TempPssSession {Import-Module DISM}
                            $HostDriversList = Invoke-Command -Session $TempPssSession {Get-WindowsDriver -Online}
                            $HostDriverReport = @()
                            ForEach($HostDriver in $HostDriversList){
                                $TempDriver = [PSCustomObject] @{
                                    'Class Description' = $HostDriver.ClassDescription
                                    'Provider Name' = $HostDriver.ProviderName
                                    'Driver Version' = $HostDriver.Version
                                    'Version Date' = $HostDriver.Date
                                }
                                $HostDriverReport += $TempDriver
                            }
                            $HostDriverReport | Table -Name 'Host Drivers' -ColumnWidths 30,30,20,20
                        }
                        #Host Roles and Features
                        Section -Style Heading4 'Roles and Features' {
                            Paragraph 'The following settings details host roles and features installed'
                            $HostRolesAndFeatures = Get-WindowsFeature -ComputerName $VMHost.Name -Credential $Credential | Where-Object {$_.Installed -eq $True}
                            [array]$HostRolesAndFeaturesReport = @()
                            ForEach($HostRoleAndFeature in $HostRolesAndFeatures){
                                $TempHostRolesAndFeaturesReport = [PSCustomObject] @{
                                    'Feature Name' = $HostRoleAndFeature.DisplayName
                                    'Feature Type' = $HostRoleAndFeature.FeatureType
                                    'Description' = $HostRoleAndFeature.Description
                                }
                            $HostRolesAndFeaturesReport += $TempHostRolesAndFeaturesReport
                            }
                            $HostRolesAndFeaturesReport | Table -Name 'Roles and Features' -ColumnWidths 20,10,70
                        }
                        #Host 3rd Party Applications
                        Section -Style Heading4 'Installed Applications' {
                            Paragraph 'The following settings details applications listed in Add/Remove Programs'
                            [array]$AddRemove = @()
                            $AddRemove += Invoke-Command -Session $TempPssSession {Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*}
                            $AddRemove += Invoke-Command -Session $TempPssSession {Get-ItemProperty HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*}
                            [array]$AddRemoveReport = @()
                            ForEach($App in $AddRemove){
                                $TempAddRemoveReport = [PSCustomObject]@{
                                    'Application Name' = $App.DisplayName
                                    'Publisher' = $App.Publisher
                                    'Version' = $App.Version
                                    'Install Date' = $App.InstallDate
                                }
                                $AddRemoveReport += $TempAddRemoveReport
                            }
                            $AddRemoveReport | Where-Object {$_.'Application Name' -notlike $null} | Sort-Object  'Application Name' | Table -Name 'Installed Applications'
                        }
                    }
                    #Local Users and Groups
                    Section -Style Heading3 'Local Users and Groups'{
                        Paragraph 'The following section details local users and groups configured'
                        Section -Style Heading4 'Local Users'{
                            Paragraph 'The following table details local users'
                            $LocalUsers = Invoke-Command -Session $TempPssSession {Get-LocalUser}
                            $LocalUsersReport = @()
                            ForEach($LocalUser in $LocalUsers){
                                $TempLocalUsersReport = [PSCustomObject]@{
                                    'User Name' = $LocalUser.Name
                                    'Description' = $LocalUser.Description
                                    'Account Enabled' = $LocalUser.Enabled
                                    'Last Logon Date' = $LocalUser.LastLogon
                                }
                                $LocalUsersReport += $TempLocalUsersReport
                            }
                            $LocalUsersReport | Table -Name 'Local Users' -ColumnWidths 20,40,10,30
                        }
                        Section -Style Heading4 'Local Groups'{
                            Paragraph 'The following table details local groups configured'
                            $LocalGroups = Invoke-Command -Session $TempPssSession {Get-LocalGroup}
                            $LocalGroupsReport = @()
                            ForEach($LocalGroup in $LocalGroups){
                                $TempLocalGroupsReport = [PSCustomObject]@{
                                    'Group Name' = $LocalGroup.Name
                                    'Description' = $LocalGroup.Description
                                }
                                $LocalGroupsReport += $TempLocalGroupsReport
                            }
                            $LocalGroupsReport | Table -Name 'Local Group Summary'
                        }
                        Section -Style Heading4 'Local Administrators'{
                            Paragraph 'The following table lists Local Administrators'
                            $LocalAdmins = Invoke-Command -Session $TempPssSession {Get-LocalGroupMember -Name 'Administrators'}
                            $LocalAdminsReport = @()
                            ForEach($LocalAdmin in $LocalAdmins){
                                $TempLocalAdminsReport = [PSCustomObject]@{
                                    'Account Name' = $LocalAdmin.Name
                                    'Account Type' = $LocalAdmin.ObjectClass
                                    'Account Source' = $LocalAdmin.PrincipalSource
                                }
                                $LocalAdminsReport += $TempLocalAdminsReport
                            }
                            $LocalAdminsReport | Table -Name 'Local Administrators'
                        }
                    }
                    #Host Firewall
                    Section -Style Heading3 'Windows Firewall'{
                        Paragraph 'The Following table is a the Windowss Firewall Summary'
                        $NetFirewallProfile = Get-NetFirewallProfile -CimSession $TempCimSession
                        $NetFirewallProfileReport = @()
                        Foreach($FirewallProfile in $NetFireWallProfile){
                            $TempNetFirewallProfileReport = [PSCustomObject]@{
                                'Profile' = $FirewallProfile.Name
                                'Profile Enabled' = $FirewallProfile.Enabled
                                'Inbound Action' = $FirewallProfile.DefaultInboundAction
                                'Outbound Action' = $FirewallProfile.DefaultOutboundAction
                            }
                            $NetFirewallProfileReport += $TempNetFirewallProfileReport
                        }
                        $NetFirewallProfileReport | Table -Name 'Windows Firewall Profiles'
                    }
                    #Host Networking
                    Section -Style Heading3 'Host Networking'{
                        Paragraph 'The following section details Host Network Configuration'
                        Section -Style Heading4 'Network Adapters'{
                            Paragraph 'The Following table details host network adapters'
                            $HostAdapters = Invoke-Command -Session $TempPssSession {Get-NetAdapter}
                            $HostAdaptersReport = @()
                            ForEach($HostAdapter in $HostAdapters){
                                $TempHostAdaptersReport = [PSCustomObject]@{
                                    'Adapter Name' = $HostAdapter.Name
                                    'Adapter Description' = $HostAdapter.InterfaceDescription
                                    'Mac Address' = $HostAdapter.MacAddress
                                    'Link Speed' = $HostAdapter.LinkSpeed
                                }
                                $HostAdaptersReport += $TempHostAdaptersReport
                            }
                            $HostAdaptersReport | Table -Name 'Network Adapters' -ColumnWidths 20,40,20,20
                            }
                        Section -Style Heading4 'IP Addresses'{
                            Paragraph 'The following table details IP Addresses assigned to hosts'
                            $NetIPs = Invoke-Command -Session $TempPssSession {Get-NetIPConfiguration | Where-Object -FilterScript {($_.NetAdapter.Status -Eq "Up")}}
                            $NetIpsReport = @()
                            ForEach($NetIp in $NetIps){
                                $TempNetIpsReport = [PSCustomObject]@{
                                    'Interface Name' = $NetIp.InterfaceAlias
                                    'Interface Description' = $NetIp.InterfaceDescription
                                    'IPv4 Addresses' = $NetIp.IPv4Address -Join ","
                                    'Subnet Mask' = $NetIp.IPv4Address[0].PrefixLength
                                    'IPv4 Gateway' = $NetIp.IPv4DefaultGateway.NextHop
                                }
                                $NetIpsReport += $TempNetIpsReport
                            }
                            $NetIpsReport | Table -Name 'Net IP Addresses'
                        }
                        Section -Style Heading4 'DNS Client'{
                            Paragraph 'The following table details the DNS Seach Domains'
                            $DnsClient = Invoke-Command -Session $TempPssSession {Get-DnsClientGlobalSetting}
                            $DnsClientReport = [PSCustomObject]@{
                                'DNS Suffix' = $DnsClient.SuffixSearchList -Join ","
                            }
                            $DnsClientReport | Table -Name "DNS Seach Domain"
                        }
                        Section -Style Heading4 'DNS Servers'{
                            Paragraph 'The following table details the DNS Server Addresses Configured'
                            $DnsServers = Invoke-Command -Session $TempPssSession {Get-DnsClientServerAddress -AddressFamily IPv4 | `
                                Where-Object {$_.ServerAddresses -notlike $null -and $_.InterfaceAlias -notlike "*isatap*"}}
                            $DnsServerReport = @()
                            ForEach($DnsServer in $DnsServers){
                                $TempDnsServerReport = [PSCustomObject]@{
                                    'Interface' = $DnsServer.InterfaceAlias
                                    'Server Address' = $DnsServer.ServerAddresses -Join ","
                            }
                            $DnsServerReport += $TempDnsServerReport
                            }
                            $DnsServerReport | Table -Name 'DNS Server Addresses' -ColumnWidths 40,60
                        }
                        $NetworkTeamCheck = Invoke-Command -Session $TempPssSession {Get-NetLbfoTeam}
                        if($NetworkTeamCheck){
                            Section -Style Heading4 'Network Team Interfaces'{
                                Paragraph 'The following table details Network Team Interfaces'
                                $NetTeams = Invoke-Command -Session $TempPssSession {Get-NetLbfoTeam}
                                $NetTeamReport = @()
                                ForEach($NetTeam in $NetTeams){
                                    $TempNetTeamReport = [PSCustomObject]@{
                                        'Team Name' = $NetTeam.Name
                                        'Team Mode' = $NetTeam.tm
                                        'Load Balancing' = $NetTeam.lba
                                        'Network Adapters' = $NetTeam.Members -Join ","
                                    }
                                    $NetTeamReport += $TempNetTeamReport
                                }
                                $NetTeamReport | Table -Name 'Network Team Interfaces'
                            }
                        }
                        Section -Style Heading4 'Network Adapter MTU'{
                            Paragraph 'The following table lists Network Adapter MTU settings'
                            $NetMtus = Invoke-Command -Session $TempPssSession {Get-NetAdapterAdvancedProperty | Where-Object {$_.DisplayName -eq 'Jumbo Packet'}}
                            $NetMtuReport = @()
                            ForEach($NetMtu in $NetMtus){
                                $TempNetMtuReport = [PSCustomObject]@{
                                    'Adapter Name' = $NetMtu.Name
                                    'MTU Size' = $NetMtu.DisplayValue
                                }
                                $NetMtuReport += $TempNetMtuReport
                            }
                            $NetMtuReport | Table -Name 'Network Adapter MTU' -ColumnWidths 50,50
                        }
                    }
                    #Host Storage
                    Section -Style Heading3 'Host Storage'{
                        Paragraph 'The following section details the storage configuration of the host'
                        #Local Disks
                        Section -Style Heading4 'Local Disks'{
                            Paragraph 'The following table details physical disks installed in the host'
                            $HostDisks = Invoke-Command -Session $TempPssSession -ScriptBlock {Get-Disk | Where-Object -FilterScript {$_.BusType -Eq "RAID" -or $_.BusType -eq "File Backed Virtual" -or $_.BusType -eq "SATA" -or $_.BusType -eq "USB"}}
                            $LocalDiskReport = @()
                            ForEach($Disk in $HostDisks){
                                $TempLocalDiskReport = [PSCustomObject]@{
                                    'Disk Number' = $Disk.Number
                                    'Model' = $Disk.Model
                                    'Serial Number' = $Disk.SerialNumber
                                    'Partition Style' = $Disk.PartitionStyle
                                    'Disk Size(GB)' = [Math]::Round($Disk.Size/1Gb)
                                }
                                $LocalDiskReport += $TempLocalDiskReport
                            }
                            $LocalDiskReport | Sort-Object -Property 'Disk Number' | Table -Name 'Local Disks'
                        }
                        #SAN Disks
                        $SanDisks = Invoke-Command -Session $TempPssSession -ScriptBlock {Get-Disk | Where-Object {$_.BusType -Eq "iSCSI"}}
                        if($SanDisks){
                            Section -Style Heading4 'SAN Disks'{
                                Paragraph 'The following section details SAN disks connected to the host'
                                $SanDiskReport = @()
                                ForEach($Disk in $SanDisks){
                                    $TempSanDiskReport = [PSCustomObject]@{
                                        'Disk Number' = $Disk.Number
                                        'Model' = $Disk.Model
                                        'Serial Number' = $Disk.SerialNumber
                                        'Partition Style' = $Disk.PartitionStyle
                                        'Disk Size(GB)' = [Math]::Round($Disk.Size/1Gb)
                                    }
                                    $SanDiskReport += $TempSanDiskReport
                                }
                                $SanDiskReport | Sort-Object -Property 'Disk Number' | Table -Name 'Local Disks'
                            }
                        }
                        #Local Volumes
                        Section -Style Heading4 'Host Volumes'{
                            Paragraph 'The following section details local volumes on the host'
                                $HostVolumes = Invoke-Command -Session $TempPssSession -ScriptBlock {Get-Volume}
                                $HostVolumeReport = @()
                                ForEach($HostVolume in $HostVolumes){
                                    $TempHostVolumeReport = [PSCustomObject]@{
                                        'Drive Letter' = $HostVolume.DriveLetter
                                        'File System Label' = $HostVolume.FileSystemLabel
                                        'File System' = $HostVolume.FileSystem
                                        'Size (GB)' = [Math]::Round($HostVolume.Size/1gb)
                                        'Free Space(GB)' = [Math]::Round($HostVolume.SizeRemaining/1gb)
                                    }
                                    $HostVolumeReport += $TempHostVolumeReport
                                }
                                $HostVolumeReport | Sort-Object 'Drive Letter' | Table -Name 'Host Volumes'
                        }
                        #iSCSI Configuration
                        $iSCSICheck = Invoke-Command -Session $TempPssSession {Get-Service -Name 'MSiSCSI'}
                        if($iSCSICheck.Status -eq 'Running'){
                            Section -Style Heading4 'Host iSCSI Settings'{
                                Paragraph 'The following section details the iSCSI configuration for the host'
                                $HostInitiator = Invoke-Command -Session $TempPssSession -ScriptBlock {Get-InitiatorPort}
                                Paragraph 'The following table details the hosts iSCI IQN'
                                $HostInitiator | Select-Object NodeAddress | Table -Name 'Host IQN'
                                Section -Style Heading5 'iSCSI Target Server'{
                                    Paragraph 'The following table details iSCSI Target Server details'
                                    $HostIscsiTargetServer = Invoke-Command -Session $TempPssSession -ScriptBlock {Get-IscsiTargetPortal}
                                    $HostIscsiTargetServer | Select-Object TargetPortalAddress,TargetPortalPortNumber | Table -Name 'iSCSI Target Servers' -ColumnWidths 50,50
                                }
                                Section -Style Heading5 'iSCIS Target Volumes'{
                                    Paragraph 'The following table details iSCSI target volumes'
                                    $HostIscsiTargetVolumes = Invoke-Command -Session $TempPssSession -ScriptBlock {Get-IscsiTarget}
                                    $HostIscsiTargetVolumeReport = @()
                                    ForEach($HostIscsiTargetVolume in $HostIscsiTargetVolumes){
                                        $TempHostIscsiTargetVolumeReport = [PSCustomObject]@{
                                            'Node Address' = $HostIscsiTargetVolume.NodeAddress
                                            'Node Connected' = $HostIscsiTargetVolume.IsConnected
                                        }
                                        $HostIscsiTargetVolumeReport += $TempHostIscsiTargetVolumeReport
                                    }
                                    $HostIscsiTargetVolumeReport | Table -Name 'iSCSI Target Volumes' -ColumnWidths 80,20
                                }
                                Section -Style Heading5 'iSCSI Connections'{
                                    Paragraph 'The following table details iSCSI Connections'
                                    $HostIscsiConnections = Invoke-Command -Session $TempPssSession -ScriptBlock {Get-IscsiConnection}
                                    $HostIscsiConnections | Select-Object ConnectionIdentifier,InitiatorAddress,TargetAddress | Table -Name 'iSCSI Connections'
                                }
                            }
                        }
                        #MPIO Configuration
                        $MPIOInstalledCheck = Invoke-Command -Session $TempPssSession {Get-WindowsFeature | Where-Object {$_.Name -like "Multipath*"}}
                        if($MPIOInstalledCheck.InstallState -eq "Installed"){
                            Section -Style Heading4 'Host MPIO Settings'{
                                Paragraph 'The following section details host MPIO Settings'
                                [string]$MpioLoadBalance = Invoke-Command -Session $TempPssSession -ScriptBlock {Get-MSDSMGlobalDefaultLoadBalancePolicy}
                                Paragraph "The default load balancing policy is: $MpioLoadBalance"
                                Section -Style Heading5 'Multipath  I/O AutoClaim'{
                                    Paragraph 'The Following table details the BUS types MPIO will automatically claim for'
                                    $MpioAutoClaim = Invoke-Command -Session $TempPssSession -ScriptBlock {Get-MSDSMAutomaticClaimSettings | Select-Object -ExpandProperty Keys}
                                    $MpioAutoClaimReport = @()
                                    foreach($key in $MpioAutoClaim){
                                        $Temp = "" | Select-Object BusType,State
                                        $Temp.BusType = $key
                                        $Temp.State = 'Enabled'
                                        $MpioAutoClaimReport += $Temp
                                    }
                                    $MpioAutoClaimReport | Table -Name 'Multipath I/O Auto Claim Settings'
                                }
                                Section -Style Heading5 'MPIO Detected Hardware'{
                                    Paragraph 'The following table details the hardware detected and claimed by MPIO'
                                    $MpioAvailableHw = Invoke-Command -Session $TempPssSession -ScriptBlock {Get-MPIOAvailableHw}
                                    $MpioAvailableHw | Select-Object VendorId,ProductId,BusType,IsMultipathed | Table -Name 'MPIO Available Hardware'
                                }
                            }
                        }
                    }
                    #HyperV Configuration
                    $HyperVInstalledCheck = Invoke-Command -Session $TempPssSession { Get-WindowsFeature | Where-Object { $_.Name -like "*Hyper-V*" } }
                    if ($HyperVInstalledCheck.InstallState -eq "Installed") {
                        Section -Style Heading4 "Hyper-V Configuration Settings" {
                            Paragraph 'The following table details the Hyper-V Server Settings'
                            $VmHost = Invoke-Command -Session $TempPssSession { Get-VMHost }
                            $VmHostReport = [PSCustomObject]@{
                                'Logical Processor Count' = $VmHost.LogicalProcessorCount
                                'Memory Capacity (GB)' = [Math]::Round($VmHost.MemoryCapacity / 1gb)
                                'VM Default Path' = $VmHost.VirtualMachinePath
                                'VM Disk Default Path' = $VmHost.VirtualHardDiskPath
                                'Supported VM Versions' = $VmHost.SupportedVmVersions -Join ","
                                'Numa Spannning Enabled' = $VmHost.NumaSpanningEnabled
                                'Iov Support' = $VmHost.IovSupport
                                'VM Migrations Enabled' = $VmHost.VirtualMachineMigrationEnabled
                                'Allow any network for Migrations' = $VmHost.UseAnyNetworkForMigrations
                                'VM Migration Authentication Type' = $VmHost.VirtualMachineMigrationAuthenticationType
                                'Max Concurrent Storage Migrations' = $VmHost.MaximumStorageMigrations
                                'Max Concurrent VM Migrations' = $VmHost.MaximumStorageMigrations
                            }
                            $VmHostReport | Table -Name 'Hyper-V Host Settings' -List -ColumnWidths 50, 50
                            Section -Style Heading5 "Hyper-V NUMA Boundaries" {
                                Paragraph 'The following table details the NUMA nodes on the host'
                                $VmHostNumaNodes = Get-VMHostNumaNode -CimSession $TempCimSession
                                [array]$VmHostNumaReport = @()
                                foreach ($Node in $VmHostNumaNodes) {
                                    $TempVmHostNumaReport = [PSCustomObject]@{
                                        'Numa Node Id' = $Node.NodeId
                                        'Memory Available(GB)' = $Node.MemoryAvailable
                                        'Memory Total(GB)' = $Node.MemoryTotal
                                    }
                                    $VmHostNumaReport += $TempVmHostNumaReport
                                }
                                $VmHostNumaReport | Table -Name 'Host NUMA Nodes'
                            }
                            Section -Style Heading5 "Hyper-V MAC Pool settings" {
                                'The following table details the Hyper-V MAC Pool'
                                $VmHostMacPool = [PSCustomObject]@{
                                'Mac Address Minimum' = $VmHost.MacAddressMinimum
                                'Mac Address Maximum' = $VmHost.MacAddressMaximum
                            }
                            $VmHostMacPool | Table -Name 'MAC Address Pool' -ColumnWidths 50, 50
                            }
                            Section -Style Heading5 "Hyper-V Management OS Adapters" {
                                Paragraph 'The following table details the Management OS Virtual Adapters created on Virtual Switches'
                                $VmOsAdapters = Get-VMNetworkAdapter -CimSession $TempCimSession -ManagementOS
                                $VmOsAdapterReport = @()
                                Foreach ($VmOsAdapter in $VmOsAdapters) {
                                    $AdapterVlan = Get-VMNetworkAdapterVlan -CimSession $TempCimSession -ManagementOS -VMNetworkAdapterName $VmOsAdapter.Name
                                    $TempVmOsAdapterReport = [PSCustomObject]@{
                                        'Name' = $VmOsAdapter.Name
                                        'Switch Name' = $VmOsAdapter.SwitchName
                                        'Mac Address' = $VmOsAdapter.MacAddress
                                        'IPv4 Address' = $VmOsAdapter.IPAddresses -Join ","
                                        'Adapter Mode' = $AdapterVlan.OperationMode
                                        'Vlan ID' = $AdapterVlan.AccessVlanId
                                    }
                                    $VmOsAdapterReport += $TempVmOsAdapterReport
                                }
                                $VmOsAdapterReport | Table -Name 'VM Management OS Adapters'
                            }
                            Section -Style Heading5 "Hyper-V vSwitch Settings" {
                                Paragraph 'The following table details the Hyper-V vSwitches configured'
                                $VmSwitches = Invoke-Command -Session $TempPssSession { Get-VMSwitch }
                                $VmSwitchesReport = @()
                                ForEach ($VmSwitch in $VmSwitches) {
                                    $TempVmSwitchesReport = [PSCustomObject]@{
                                        'Switch Name' = $VmSwitch.Name
                                        'Switch Type' = $VmSwitch.SwitchType
                                        'Embedded Team' = $VmSwitch.EmbeddedTeamingEnabled
                                        'Interface Description' = $VmSwitch.NetAdapterInterfaceDescription
                                    }
                                    $VmSwitchesReport += $TempVmSwitchesReport
                                }
                                $VmSwitchesReport | Table -Name 'Virtual Switch Summary' -ColumnWidths 40, 10, 10, 40
                                Foreach ($VmSwitch in $VmSwitches) {
                                    Section -Style Heading6 ($VmSwitch.Name) {
                                        Paragraph 'The following table details the Hyper-V vSwitch'
                                        $VmSwitchReport = [PSCustomObject]@{
                                            'Switch Name' = $VmSwitch.Name
                                            'Switch Type' = $VmSwitch.SwitchType
                                            'Switch Embedded Teaming Status' = $VmSwitch.EmbeddedTeamingEnabled
                                            'Bandwidth Reservation Mode' = $VmSwitch.BandwidthReservationMode
                                            'Bandwidth Reservation Percentage' = $VmSwitch.Percentage
                                            'Management OS Allowed' = $VmSwitch.AllowManagementOS
                                            'Physical Adapters' = $VmSwitch.NetAdapterInterfaceDescriptions -Join ","
                                            'IOV Support' = $VmSwitch.IovSupport
                                            'IOV Support Reasons' = $VmSwitch.IovSupportReasons
                                            'Available VM Queues' = $VmSwitch.AvailableVMQueues
                                            'Packet Direct Enabled' = $VmSwitch.PacketDirectinUse
                                        }
                                        $VmSwitchReport | Table -Name 'VM Switch Details' -List -ColumnWidths 50, 50
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Remove-PSSession $TempPssSession
            Remove-CimSession $TempCimSession
        }

	}
	#endregion foreach loop
}
