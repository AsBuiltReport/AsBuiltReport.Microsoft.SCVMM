<p align="center">
    <a href="https://www.asbuiltreport.com/" alt="AsBuiltReport"></a>
            <img src='https://avatars.githubusercontent.com/u/42958564' width="8%" height="8%" /></a>
</p>
<p align="center">
    <a href="https://www.powershellgallery.com/packages/AsBuiltReport.Microsoft.SCVMM/" alt="PowerShell Gallery Version">
        <img src="https://img.shields.io/powershellgallery/v/AsBuiltReport.Microsoft.SCVMM.svg" /></a>
    <a href="https://www.powershellgallery.com/packages/AsBuiltReport.Microsoft.SCVMM/" alt="PS Gallery Downloads">
        <img src="https://img.shields.io/powershellgallery/dt/AsBuiltReport.Microsoft.SCVMM.svg" /></a>
    <a href="https://www.powershellgallery.com/packages/AsBuiltReport.Microsoft.SCVMM/" alt="PS Platform">
        <img src="https://img.shields.io/powershellgallery/p/AsBuiltReport.Microsoft.SCVMM.svg" /></a>
</p>
<p align="center">
    <a href="https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM/graphs/commit-activity" alt="GitHub Last Commit">
        <img src="https://img.shields.io/github/last-commit/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM/master.svg" /></a>
    <a href="https://raw.githubusercontent.com/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM/master/LICENSE" alt="GitHub License">
        <img src="https://img.shields.io/github/license/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM.svg" /></a>
    <a href="https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM/graphs/contributors" alt="GitHub Contributors">
        <img src="https://img.shields.io/github/contributors/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM.svg"/></a>
</p>
<p align="center">
    <a href="https://twitter.com/AsBuiltReport" alt="Twitter">
            <img src="https://img.shields.io/twitter/follow/AsBuiltReport.svg?style=social"/></a>
</p>
<p align="center">
    <a href='https://ko-fi.com/B0B7DDGZ7' target='_blank'><img height='36' style='border:0px;height:36px;' src='https://cdn.ko-fi.com/cdn/kofi1.png?v=3' border='0' alt='Buy Me a Coffee at ko-fi.com' /></a>
</p>

# Microsoft SCVMM As Built Report

Microsoft SCVMM As Built Report is a PowerShell module which works in conjunction with [AsBuiltReport.Core](https://github.com/AsBuiltReport/AsBuiltReport.Core).

[AsBuiltReport](https://github.com/AsBuiltReport/AsBuiltReport) is an open-sourced community project which utilizes PowerShell to produce as-built documentation in multiple document formats for multiple vendors and technologies.

Please refer to the AsBuiltReport [website](https://www.asbuiltreport.com) for more detailed information about this project.

# :books: Sample Reports

## Sample Report - Custom Style 1

Sample Microsoft SCVMM As Built report HTML file: [Sample Microsoft SCVMM As-Built Report.html](https://htmlpreview.github.io/?https://raw.githubusercontent.com/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM/dev/Samples/Sample%20Microsoft%20SCVMM%20As%20Built%20Report.html "Sample Microsoft SCVMM As-Built Report")

# :beginner: Getting Started

Below are the instructions on how to install, configure and generate a Microsoft SCVMM As Built report.

## :floppy_disk: Supported Versions
<!-- ********** Update supported SCVMM versions ********** -->
The Microsoft SCVMM As Built Report supports the following SCVMM Server versions;

- 2016, 2019, 2022 & 2025

### PowerShell

This report is compatible with the following PowerShell versions;

<!-- ********** Update supported PowerShell versions ********** -->
| Windows PowerShell 5.1 |    PowerShell 7    |
| :--------------------: | :----------------: |
|   :white_check_mark:   | :x: |

## :wrench: System Requirements
<!-- ********** Update system requirements ********** -->
PowerShell 5.1 and the following PowerShell modules are required for generating a Microsoft SCVMM As Built report.

- [AsBuiltReport.Microsoft.SCVMM Module](https://www.powershellgallery.com/packages/AsBuiltReport.Microsoft.SCVMM/)
- [Hyper-V Module](https://docs.microsoft.com/en-us/powershell/module/hyper-v/?view=windowsserver2022-ps)
- [FailoverClusters Module](https://learn.microsoft.com/en-us/powershell/module/failoverclusters/?view=windowsserver2022-ps)
- [Diagrammer.Core Module](https://www.powershellgallery.com/packages/Diagrammer.Core/)


### Linux & macOS

This report does not support Linux or Mac due to the fact that the SCVMM modules are dependent on the .NET Framework. Until Microsoft migrates these modules to native PowerShell Core, only PowerShell >= 5.1.x will be supported on SCVMM.

### :closed_lock_with_key: Required Privileges

A Microsoft SCVMM As Built Report can be generated with Administrator level role. Since this report relies extensively on the WinRM component, you should make sure that it is enabled and configured

## :package: Module Installation

The installation of the modules will depend on the roles that are being served on the server to be documented.

### PowerShell v5.x running on a Windows server (VMM server)

```powershell
Install-Module AsBuiltReport.Microsoft.SCVMM

# Hyper-V Server powershell modules
Install-WindowsFeature -Name Hyper-V-PowerShell

# FailOver Cluster powershell modules
Install-WindowsFeature -Name RSAT-Clustering-PowerShell
```

### GitHub

If you are unable to use the PowerShell Gallery, you can still install the module manually. Ensure you repeat the following steps for the [system requirements](https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM#wrench-system-requirements) also.

1. Download the code package / [latest release](https://github.com/AsBuiltReport/AsBuiltReport.Microsoft.SCVMM/releases/latest) zip from GitHub
2. Extract the zip file
3. Copy the folder `AsBuiltReport.Microsoft.SCVMM` to a path that is set in `$env:PSModulePath`.
4. Open a PowerShell terminal window and unblock the downloaded files with

    ```powershell
    $path = (Get-Module -Name AsBuiltReport.Microsoft.SCVMM -ListAvailable).ModuleBase; Unblock-File -Path $path\*.psd1; Unblock-File -Path $path\Src\Public\*.ps1; Unblock-File -Path $path\Src\Private\*.ps1
    ```

5. Close and reopen the PowerShell terminal window.

_Note: You are not limited to installing the module to those example paths, you can add a new entry to the environment variable PSModulePath if you want to use another path._

## :pencil2: Configuration

The Microsoft SCVMM As Built Report utilizes a JSON file to allow configuration of report information, options, detail and healthchecks.

A Microsoft SCVMM report configuration file can be generated by executing the following command;

```powershell
New-AsBuiltReportConfig -Report Microsoft.SCVMM -FolderPath <User specified folder> -Filename <Optional>
```

Executing this command will copy the default Microsoft Windows report JSON configuration to a user specified folder.

All report settings can then be configured via the JSON file.

The following provides information of how to configure each schema within the report's JSON file.

### Report

The **Report** schema provides configuration of the Microsoft SCVMM report information.

| Sub-Schema          | Setting      | Default                         | Description                                                  |
| ------------------- | ------------ | ------------------------------- | ------------------------------------------------------------ |
| Name                | User defined | Microsoft SCVMM As Built Report | The name of the As Built Report                              |
| Version             | User defined | 1.0                             | The report version                                           |
| Status              | User defined | Released                        | The report release status                                    |
| ShowCoverPageImage  | true / false | true                            | Toggle to enable/disable the display of the cover page image |
| ShowTableOfContents | true / false | true                            | Toggle to enable/disable table of contents                   |
| ShowHeaderFooter    | true / false | true                            | Toggle to enable/disable document headers & footers          |
| ShowTableCaptions   | true / false | true                            | Toggle to enable/disable table captions/numbering            |

### Options

The **Options** schema allows certain options within the report to be toggled on or off.

| Sub-Schema             | Setting      | Default | Description                                                                   |
| ---------------------- | ------------ | ------- | ----------------------------------------------------------------------------- |
| DiagramColumnSize      | int          | 3       | Set the diagram node table size                                               |
| DiagramTheme           | string       | White   | Set the diagram theme (Black/White/Neon)                                      |
| DiagramWaterMark       | string       | empty   | Set the diagram watermark                                                     |
| DiagramType            | true / false | true    | Toggle to enable/disable the export of individual diagram diagrams            |
| EnableDiagrams         | true / false | false   | Toggle to enable/disable infrastructure diagrams                              |
| EnableDiagramsDebug    | true / false | false   | Toggle to enable/disable diagram debug option                                 |
| EnableDiagramSignature | true / false | false   | Toggle to enable/disable diagram signature (bottom right corner)              |
| ExportDiagrams         | true / false | true    | Toggle to enable/disable diagram export option                                |
| ExportDiagramsFormat   | string array | pdf     | Set the format used to export the infrastructure diagram (dot, png, pdf, svg) |
| SignatureAuthorName    | string       | empty   | Set the signature author name                                                 |
| SignatureCompanyName   | string       | empty   | Set the signature company name                                                |

### InfoLevel

The **InfoLevel** schema allows configuration of each section of the report at a granular level. The following sections can be set.

There are 3 levels (0-2) of detail granularity for each section as follows;

| Setting | InfoLevel   | Description                                                          |
| :-----: | ----------- | -------------------------------------------------------------------- |
|    0    | Disabled    | Does not collect or display any information                          |
|    1    | Enabled     | Provides summarized information for a collection of objects          |
|    2    | Adv Summary | Provides condensed, detailed information for a collection of objects |

The table below outlines the default and maximum **InfoLevel** settings for each section.

| Sub-Schema       | Default Setting | Maximum Setting |
| ---------------- | :-------------: | :-------------: |
| Clusters         |        1        |        2        |
| Hosts            |        1        |        1        |
| HostGroups       |        1        |        1        |
| Infrastructures  |        1        |        2        |
| LibraryTemplates |        1        |        2        |
| Networking       |        1        |        2        |

### Healthcheck

The **Healthcheck** schema is used to toggle health checks on or off.

## :computer: Examples

There are a few examples listed below on running the AsBuiltReport script against a Microsoft Windows server target. Refer to the `README.md` file in the main AsBuiltReport project repository for more examples.

```powershell

# Generate a Microsoft SCVMM As Built Report for Server 'scvmm-server-01v.contoso.local' using specified credentials. Export report to HTML & DOCX formats. Use default report style. Append timestamp to report filename. Save reports to 'C:\Users\Jon\Documents'
PS C:\> New-AsBuiltReport -Report Microsoft.SCVMM -Target 'scvmm-server-01v.contoso.local' -Username 'administrator@contoso.local' -Password 'P@ssw0rd' -Format Html,Word -OutputFolderPath 'C:\Users\Jon\Documents' -Timestamp

# Generate a Microsoft SCVMM As Built Report for Server 'scvmm-server-01v.contoso.local' using specified credentials and report configuration file. Export report to Text, HTML & DOCX formats. Use default report style. Save reports to 'C:\Users\Jon\Documents'. Display verbose messages to the console.
PS C:\> New-AsBuiltReport -Report Microsoft.SCVMM -Target 'scvmm-server-01v.contoso.local' -Username 'administrator@contoso.local' -Password 'P@ssw0rd' -Format Text,Html,Word -OutputFolderPath 'C:\Users\Jon\Documents' -ReportConfigFilePath 'C:\Users\Jon\AsBuiltReport\AsBuiltReport.Microsoft.SCVMM.json' -Verbose

# Generate a Microsoft SCVMM As Built Report for Server 'scvmm-server-01v.contoso.local' using stored credentials. Export report to HTML & Text formats. Use default report style. Highlight environment issues within the report. Save reports to 'C:\Users\Jon\Documents'.
PS C:\> $Creds = Get-Credential
PS C:\> New-AsBuiltReport -Report Microsoft.SCVMM -Target 'scvmm-server-01v.contoso.local' -Credential $Creds -Format Html,Text -OutputFolderPath 'C:\Users\Jon\Documents' -EnableHealthCheck

# Generate a Microsoft SCVMM As Built Report for Server 'scvmm-server-01v.contoso.local' using specified credentials. Export report to HTML & DOCX formats. Use default report style. Reports are saved to the user profile folder by default. Attach and send reports via e-mail.
PS C:\> New-AsBuiltReport -Report Microsoft.SCVMM -Target 'scvmm-server-01v.contoso.local' -Username 'administrator@contoso.local' -Password 'P@ssw0rd' -Format Html,Word -OutputFolderPath 'C:\Users\Jon\Documents' -SendEmail
```

## :x: Known Issues

- Issues with WinRM when using the IP address instead of the "Fully Qualified Domain Name".