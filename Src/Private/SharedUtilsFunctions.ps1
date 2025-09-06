function ConvertTo-TextYN {
    <#
    .SYNOPSIS
        Used by As Built Report to convert true or false automatically to Yes or No.
    .DESCRIPTION

    .NOTES
        Version:        0.3.0
        Author:         LEE DAILEY

    .EXAMPLE

    .LINK

    #>
    [CmdletBinding()]
    [OutputType([String])]
    param (
        [Parameter (
            Position = 0,
            Mandatory)]
        [AllowEmptyString()]
        [string] $TEXT
    )

    switch ($TEXT) {
        "" { "--"; break }
        " " { "--"; break }
        $Null { "--"; break }
        "True" { "Yes"; break }
        "False" { "No"; break }
        default { $TEXT }
    }
} # end


function ConvertTo-FileSizeString {
    <#
    .SYNOPSIS
    Used by As Built Report to convert bytes automatically to GB or TB based on size.
    .DESCRIPTION
    .NOTES
        Version:        0.1.0
        Author:         Jonathan Colon
    .EXAMPLE
    .LINK
    #>
    [CmdletBinding()]
    [OutputType([String])]
    param
    (
        [Parameter (
            Position = 0,
            Mandatory)]
        [int64] $Size,
        [Parameter(
            Position = 1,
            Mandatory = $false,
            HelpMessage = 'Please provide the source space unit'
        )]
        [ValidateSet('MB', 'GB', 'TB', 'PB')]
        [string] $SourceSpaceUnit,
        [Parameter(
            Position = 2,
            Mandatory = $false,
            HelpMessage = 'Please provide the space unit to output'
        )]
        [ValidateSet('MB', 'GB', 'TB', 'PB')]
        [string] $TargetSpaceUnit,
        [Parameter(
            Position = 3,
            Mandatory = $false,
            HelpMessage = 'Please provide the value to round the storage unit'
        )]
        [int] $RoundUnits = 0
    )

    if ($SourceSpaceUnit) {
        return "$([math]::Round(($Size * $("1" + $SourceSpaceUnit) / $("1" + $TargetSpaceUnit)), $RoundUnits)) $TargetSpaceUnit"
    } else {
        $Unit = switch ($Size) {
            { $Size -gt 1PB } { 'PB' ; break }
            { $Size -gt 1TB } { 'TB' ; break }
            { $Size -gt 1GB } { 'GB' ; break }
            { $Size -gt 1Mb } { 'MB' ; break }
            Default { 'KB' }
        }
        return "$([math]::Round(($Size / $("1" + $Unit)), $RoundUnits)) $Unit"
    }
} # end
function Convert-Size {
    [cmdletbinding()]
    param(
        [validateset("Bytes", "KB", "MB", "GB", "TB")]
        [string]$From,
        [validateset("Bytes", "KB", "MB", "GB", "TB")]
        [string]$To,
        [Parameter(Mandatory = $true)]
        [double]$Value,
        [int]$Precision = 4
    )
    switch ($From) {
        "Bytes" { $value = $Value }
        "KB" { $value = $Value * 1024 }
        "MB" { $value = $Value * 1024 * 1024 }
        "GB" { $value = $Value * 1024 * 1024 * 1024 }
        "TB" { $value = $Value * 1024 * 1024 * 1024 * 1024 }
    }

    switch ($To) {
        "Bytes" { return $value }
        "KB" { $Value = $Value / 1KB }
        "MB" { $Value = $Value / 1MB }
        "GB" { $Value = $Value / 1GB }
        "TB" { $Value = $Value / 1TB }

    }

    return [Math]::Round($value, $Precision, [MidPointRounding]::AwayFromZero)
}

function Get-PieChart {
    <#
    .SYNOPSIS
    Used by As Built Report to generate PScriboChart pie charts.
    .DESCRIPTION
    .NOTES
        Version:        0.1.0
        Author:         Jonathan Colon
    .EXAMPLE
    .LINK
    #>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter (
            Position = 0,
            Mandatory)]
        [System.Array]
        $SampleData,
        [String]
        $ChartName,
        [String]
        $XField,
        [String]
        $YField,
        [String]
        $ChartLegendName,
        [String]
        $ChartLegendAlignment = 'Center',
        [String]
        $ChartTitleName = ' ',
        [String]
        $ChartTitleText = ' ',
        [int]
        $Width = 600,
        [int]
        $Height = 400,
        [Switch]
        $Status,
        [bool]
        $ReversePalette = $false
    )

    $StatusCustomPalette = @(
        [System.Drawing.ColorTranslator]::FromHtml('#DFF0D0')
        [System.Drawing.ColorTranslator]::FromHtml('#FFF4C7')
        [System.Drawing.ColorTranslator]::FromHtml('#FEDDD7')
        [System.Drawing.ColorTranslator]::FromHtml('#878787')
    )

    $AbrCustomPalette = @(
        [System.Drawing.ColorTranslator]::FromHtml('#d5e2ff')
        [System.Drawing.ColorTranslator]::FromHtml('#bbc9e9')
        [System.Drawing.ColorTranslator]::FromHtml('#a2b1d3')
        [System.Drawing.ColorTranslator]::FromHtml('#8999bd')
        [System.Drawing.ColorTranslator]::FromHtml('#7082a8')
        [System.Drawing.ColorTranslator]::FromHtml('#586c93')
        [System.Drawing.ColorTranslator]::FromHtml('#40567f')
        [System.Drawing.ColorTranslator]::FromHtml('#27416b')
        [System.Drawing.ColorTranslator]::FromHtml('#072e58')
    )

    $VeeamCustomPalette = @(
        [System.Drawing.ColorTranslator]::FromHtml('#ddf6ed')
        [System.Drawing.ColorTranslator]::FromHtml('#c3e2d7')
        [System.Drawing.ColorTranslator]::FromHtml('#aacec2')
        [System.Drawing.ColorTranslator]::FromHtml('#90bbad')
        [System.Drawing.ColorTranslator]::FromHtml('#77a898')
        [System.Drawing.ColorTranslator]::FromHtml('#5e9584')
        [System.Drawing.ColorTranslator]::FromHtml('#458370')
        [System.Drawing.ColorTranslator]::FromHtml('#2a715d')
        [System.Drawing.ColorTranslator]::FromHtml('#005f4b')
    )

    if ($Options.ReportStyle -eq "Veeam") {
        $BorderColor = 'DarkGreen'
    } else {
        $BorderColor = 'DarkBlue'
    }

    $exampleChart = New-Chart -Name $ChartName -Width $Width -Height $Height -BorderStyle Dash -BorderWidth 1 -BorderColor $BorderColor

    $addChartAreaParams = @{
        Chart = $exampleChart
        Name = 'exampleChartArea'
        AxisXInterval = 1
    }
    $exampleChartArea = Add-ChartArea @addChartAreaParams -PassThru

    if ($Status) {
        $CustomPalette = $StatusCustomPalette
    } elseif ($Options.ReportStyle -eq 'Veeam') {
        $CustomPalette = $VeeamCustomPalette

    } else {
        $CustomPalette = $AbrCustomPalette
    }

    $addChartSeriesParams = @{
        Chart = $exampleChart
        ChartArea = $exampleChartArea
        Name = 'exampleChartSeries'
        XField = $XField
        YField = $YField
        CustomPalette = $CustomPalette
        ColorPerDataPoint = $true
        ReversePalette = $ReversePalette
    }

    $sampleData | Add-PieChartSeries @addChartSeriesParams

    $addChartLegendParams = @{
        Chart = $exampleChart
        Name = $ChartLegendName
        TitleAlignment = $ChartLegendAlignment
    }
    Add-ChartLegend @addChartLegendParams

    $addChartTitleParams = @{
        Chart = $exampleChart
        ChartArea = $exampleChartArea
        Name = $ChartTitleName
        Text = $ChartTitleText
        Font = New-Object -TypeName 'System.Drawing.Font' -ArgumentList @('Segoe Ui', '12', [System.Drawing.FontStyle]::Bold)
    }
    Add-ChartTitle @addChartTitleParams

    $TempPath = Resolve-Path ([System.IO.Path]::GetTempPath())

    $ChartImage = Export-Chart -Chart $exampleChart -Path $TempPath.Path -Format "PNG" -PassThru

    $Base64Image = [convert]::ToBase64String((Get-Content $ChartImage -Encoding byte))

    Remove-Item -Path $ChartImage.FullName

    return $Base64Image

} # end

function Get-ColumnChart {
    <#
    .SYNOPSIS
    Used by As Built Report to generate PScriboChart column charts.
    .DESCRIPTION
    .NOTES
        Version:        0.1.0
        Author:         Jonathan Colon
    .EXAMPLE
    .LINK
    #>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter (
            Position = 0,
            Mandatory)]
        [System.Array]
        $SampleData,
        [String]
        $ChartName,
        [String]
        $AxisXTitle,
        [String]
        $AxisYTitle,
        [String]
        $XField,
        [String]
        $YField,
        [String]
        $ChartAreaName,
        [String]
        $ChartTitleName = ' ',
        [String]
        $ChartTitleText = ' ',
        [int]
        $Width = 600,
        [int]
        $Height = 400,
        [Switch]
        $Status,
        [bool]
        $ReversePalette = $false
    )

    $StatusCustomPalette = @(
        [System.Drawing.ColorTranslator]::FromHtml('#DFF0D0')
        [System.Drawing.ColorTranslator]::FromHtml('#FFF4C7')
        [System.Drawing.ColorTranslator]::FromHtml('#FEDDD7')
        [System.Drawing.ColorTranslator]::FromHtml('#878787')
    )

    $AbrCustomPalette = @(
        [System.Drawing.ColorTranslator]::FromHtml('#d5e2ff')
        [System.Drawing.ColorTranslator]::FromHtml('#bbc9e9')
        [System.Drawing.ColorTranslator]::FromHtml('#a2b1d3')
        [System.Drawing.ColorTranslator]::FromHtml('#8999bd')
        [System.Drawing.ColorTranslator]::FromHtml('#7082a8')
        [System.Drawing.ColorTranslator]::FromHtml('#586c93')
        [System.Drawing.ColorTranslator]::FromHtml('#40567f')
        [System.Drawing.ColorTranslator]::FromHtml('#27416b')
        [System.Drawing.ColorTranslator]::FromHtml('#072e58')
    )

    $VeeamCustomPalette = @(
        [System.Drawing.ColorTranslator]::FromHtml('#ddf6ed')
        [System.Drawing.ColorTranslator]::FromHtml('#c3e2d7')
        [System.Drawing.ColorTranslator]::FromHtml('#aacec2')
        [System.Drawing.ColorTranslator]::FromHtml('#90bbad')
        [System.Drawing.ColorTranslator]::FromHtml('#77a898')
        [System.Drawing.ColorTranslator]::FromHtml('#5e9584')
        [System.Drawing.ColorTranslator]::FromHtml('#458370')
        [System.Drawing.ColorTranslator]::FromHtml('#2a715d')
        [System.Drawing.ColorTranslator]::FromHtml('#005f4b')
    )

    if ($Options.ReportStyle -eq "Veeam") {
        $BorderColor = 'DarkGreen'
    } else {
        $BorderColor = 'DarkBlue'
    }

    $exampleChart = New-Chart -Name $ChartName -Width $Width -Height $Height -BorderStyle Dash -BorderWidth 1 -BorderColor $BorderColor

    $addChartAreaParams = @{
        Chart = $exampleChart
        Name = $ChartAreaName
        AxisXTitle = $AxisXTitle
        AxisYTitle = $AxisYTitle
        NoAxisXMajorGridLines = $true
        NoAxisYMajorGridLines = $true
        AxisXInterval = 1
    }
    $exampleChartArea = Add-ChartArea @addChartAreaParams -PassThru

    if ($Status) {
        $CustomPalette = $StatusCustomPalette
    } elseif ($Options.ReportStyle -eq 'Veeam') {
        $CustomPalette = $VeeamCustomPalette

    } else {
        $CustomPalette = $AbrCustomPalette
    }

    $addChartSeriesParams = @{
        Chart = $exampleChart
        ChartArea = $exampleChartArea
        Name = 'exampleChartSeries'
        XField = $XField
        YField = $YField
        CustomPalette = $CustomPalette
        ColorPerDataPoint = $true
        ReversePalette = $ReversePalette
    }

    $sampleData | Add-ColumnChartSeries @addChartSeriesParams

    $addChartTitleParams = @{
        Chart = $exampleChart
        ChartArea = $exampleChartArea
        Name = $ChartTitleName
        Text = $ChartTitleText
        Font = New-Object -TypeName 'System.Drawing.Font' -ArgumentList @('Segoe Ui', '12', [System.Drawing.FontStyle]::Bold)
    }
    Add-ChartTitle @addChartTitleParams

    $TempPath = Resolve-Path ([System.IO.Path]::GetTempPath())

    $ChartImage = Export-Chart -Chart $exampleChart -Path $TempPath.Path -Format "PNG" -PassThru

    if ($PassThru) {
        Write-Output -InputObject $chartFileItem
    }

    $Base64Image = [convert]::ToBase64String((Get-Content $ChartImage -Encoding byte))

    Remove-Item -Path $ChartImage.FullName

    return $Base64Image

} # end

function Get-WindowsTimePeriod {
    <#
    .SYNOPSIS
    Used by As Built Report to generate time period table.
    .DESCRIPTION
    .NOTES
        Version:        0.1.0
        Author:         Jonathan Colon
    .EXAMPLE
    .LINK
    #>
    [CmdletBinding()]
    param
    (
        [Parameter (
            Position = 0,
            Mandatory)]
        [System.Array]
        $InputTimePeriod
    )

    $OutObj = @()
    $Hours24 = [ordered]@{
        0 = 12
        1 = 1
        2 = 2
        3 = 3
        4 = 4
        5 = 5
        6 = 6
        7 = 7
        8 = 8
        9 = 9
        10 = 10
        11 = 11
        12 = 12
        13 = 1
        14 = 2
        15 = 3
        16 = 4
        17 = 5
        18 = 6
        19 = 7
        20 = 8
        21 = 9
        22 = 10
        23 = 11
    }
    $ScheduleTimePeriod = $InputTimePeriod -split '(.{48})' | Where-Object { $_ }

    foreach ($OBJ in $Hours24.GetEnumerator()) {

        $inObj = [ordered] @{
            'H' = $OBJ.Value
            'Sun' = $ScheduleTimePeriod[0].Split(',')[$OBJ.Key]
            'Mon' = $ScheduleTimePeriod[1].Split(',')[$OBJ.Key]
            'Tue' = $ScheduleTimePeriod[2].Split(',')[$OBJ.Key]
            'Wed' = $ScheduleTimePeriod[3].Split(',')[$OBJ.Key]
            'Thu' = $ScheduleTimePeriod[4].Split(',')[$OBJ.Key]
            'Fri' = $ScheduleTimePeriod[5].Split(',')[$OBJ.Key]
            'Sat' = $ScheduleTimePeriod[6].Split(',')[$OBJ.Key]
        }
        $OutObj += $inobj
    }

    return $OutObj

} # end

# Variable translating Icon to Image Path ($IconPath)
$script:Images = @{
}

function ConvertTo-HashToYN {
    <#
    .SYNOPSIS
        Used by As Built Report to convert array content true or false automatically to Yes or No.
    .DESCRIPTION

    .NOTES
        Version:        0.2.0
        Author:         Jonathan Colon

    .EXAMPLE

    .LINK

    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param (
        [Parameter (Position = 0, Mandatory)]
        [AllowEmptyString()]
        [System.Collections.Specialized.OrderedDictionary] $TEXT
    )

    $result = [ordered] @{}
    foreach ($i in $TEXT.GetEnumerator()) {
        try {
            $result.add($i.Key, (ConvertTo-TextYN $i.Value))
        } catch {
            $result.add($i.Key, ($i.Value))
        }
    }
    if ($result) {
        return $result
    } else { return $TEXT }
} # end