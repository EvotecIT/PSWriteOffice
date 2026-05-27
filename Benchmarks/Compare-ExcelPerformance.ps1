param(
    [ValidateSet('Smoke', 'Standard', 'Large', 'Full', 'SuperLarge')]
    [string] $Suite = 'Standard',

    [object[]] $RowCount,

    [int] $RepeatCount = 0,

    [string[]] $Scenario,

    [string[]] $Engine = @('PSWriteOffice', 'ImportExcel', 'ExcelFast'),

    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\Ignore\Benchmarks\ExcelPerformance'),

    [switch] $ListScenarios,

    [switch] $SkipFollowUps,

    [switch] $SkipImportExcelInstall,

    [switch] $SkipExcelFastInstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$invariantCulture = [Globalization.CultureInfo]::InvariantCulture
[Threading.Thread]::CurrentThread.CurrentCulture = $invariantCulture
[Threading.Thread]::CurrentThread.CurrentUICulture = $invariantCulture

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$moduleRoot = Join-Path $OutputDirectory 'Modules'
$workRoot = Join-Path $OutputDirectory ('Run-{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $PID)
$validEngines = @('PSWriteOffice', 'ImportExcel', 'ExcelFast')

function Resolve-EngineList {
    param([string[]] $Value)

    $resolved = [Collections.Generic.List[string]]::new()
    foreach ($item in @($Value)) {
        foreach ($engineName in ($item -split ',')) {
            $name = $engineName.Trim()
            if ([string]::IsNullOrWhiteSpace($name)) {
                continue
            }

            $match = @($validEngines | Where-Object { $_ -eq $name })
            if ($match.Count -eq 0) {
                throw "Unknown engine '$name'. Valid engines: $($validEngines -join ', ')."
            }

            if (-not $resolved.Contains($match[0])) {
                $resolved.Add($match[0])
            }
        }
    }

    if ($resolved.Count -eq 0) {
        throw "At least one engine is required. Valid engines: $($validEngines -join ', ')."
    }

    , $resolved.ToArray()
}

$Engine = Resolve-EngineList -Value $Engine

function Resolve-RowCountList {
    param([object[]] $Value)

    $resolved = [Collections.Generic.List[int]]::new()
    foreach ($item in @($Value)) {
        foreach ($rowCountText in ($item -split ',')) {
            $text = $rowCountText.Trim()
            if ([string]::IsNullOrWhiteSpace($text)) {
                continue
            }

            try {
                $rowCountValue = [int]::Parse($text, [Globalization.NumberStyles]::None, $invariantCulture)
            } catch {
                throw "Invalid row count '$text'. Use plain integers such as 10000, not grouped numbers."
            }

            if ($rowCountValue -le 0) {
                throw "Invalid row count '$text'. Row counts must be greater than zero."
            }

            $resolved.Add($rowCountValue)
        }
    }

    if ($resolved.Count -eq 0) {
        throw 'At least one row count is required.'
    }

    , $resolved.ToArray()
}

function Resolve-StringList {
    param([string[]] $Value)

    $resolved = [Collections.Generic.List[string]]::new()
    foreach ($item in @($Value)) {
        foreach ($textValue in ($item -split ',')) {
            $text = $textValue.Trim()
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $resolved.Add($text)
            }
        }
    }

    , $resolved.ToArray()
}

if ($Scenario -and $Scenario.Count -gt 0) {
    $Scenario = Resolve-StringList -Value $Scenario
}

function Add-ModulePath {
    param([string] $Path)

    if (-not ($env:PSModulePath -split [IO.Path]::PathSeparator | Where-Object { $_ -eq $Path })) {
        $env:PSModulePath = $Path + [IO.Path]::PathSeparator + $env:PSModulePath
    }
}

function Ensure-ImportExcel {
    if (Get-Module -ListAvailable ImportExcel | Sort-Object Version -Descending | Select-Object -First 1) {
        return
    }

    Add-ModulePath -Path $moduleRoot
    if (Get-Module -ListAvailable ImportExcel | Sort-Object Version -Descending | Select-Object -First 1) {
        return
    }

    if ($SkipImportExcelInstall.IsPresent) {
        throw 'ImportExcel is not installed. Rerun without -SkipImportExcelInstall to save it under the benchmark module folder.'
    }

    Save-Module -Name ImportExcel -Path $moduleRoot -Repository PSGallery -Force
    Add-ModulePath -Path $moduleRoot
}

function Ensure-ExcelFast {
    if (Get-Module -ListAvailable ExcelFast | Sort-Object Version -Descending | Select-Object -First 1) {
        return
    }

    Add-ModulePath -Path $moduleRoot
    if (Get-Module -ListAvailable ExcelFast | Sort-Object Version -Descending | Select-Object -First 1) {
        return
    }

    if ($SkipExcelFastInstall.IsPresent) {
        throw 'ExcelFast is not installed. Rerun without -SkipExcelFastInstall to save it under the benchmark module folder.'
    }

    Save-Module -Name ExcelFast -Path $moduleRoot -Repository PSGallery -AllowPrerelease -Force
    Add-ModulePath -Path $moduleRoot
}

function New-BenchmarkRows {
    param([int] $Count)

    for ($i = 1; $i -le $Count; $i++) {
        [pscustomobject]@{
            Id          = $i
            Name        = 'Server-{0:000000}' -f $i
            Department  = 'Department-{0}' -f ($i % 25)
            Region      = @('NA', 'EU', 'APAC', 'LATAM')[$i % 4]
            IsEnabled   = ($i % 3) -ne 0
            Created     = ([datetime]'2024-01-01').AddMinutes($i)
            Score       = [math]::Round(($i * 1.137) % 1000, 3)
            Owner       = 'owner{0}@example.test' -f ($i % 250)
            TicketCount = $i % 17
            Notes       = 'Benchmark row {0}' -f $i
        }
    }
}

function New-WideBenchmarkRows {
    param([int] $Count)

    for ($i = 1; $i -le $Count; $i++) {
        $row = [ordered]@{
            Id      = $i
            Name    = 'Wide-{0:000000}' -f $i
            Created = ([datetime]'2024-01-01').AddSeconds($i)
            Enabled = ($i % 2) -eq 0
        }

        for ($column = 1; $column -le 36; $column++) {
            $row["Metric$column"] = [math]::Round((($i + $column) * 1.017) % 10000, 4)
        }

        [pscustomobject]$row
    }
}

function New-BenchmarkDataTable {
    param([int] $Count)

    $table = [Data.DataTable]::new('Data')
    $null = $table.Columns.Add('Id', [int])
    $null = $table.Columns.Add('Name', [string])
    $null = $table.Columns.Add('Created', [datetime])
    $null = $table.Columns.Add('Amount', [decimal])
    $null = $table.Columns.Add('Enabled', [bool])
    $null = $table.Columns.Add('Notes', [string])

    for ($i = 1; $i -le $Count; $i++) {
        $row = $table.NewRow()
        $row.Id = $i
        $row.Name = 'Account-{0:000000}' -f $i
        $row.Created = ([datetime]'2024-01-01').AddMinutes($i)
        $row.Amount = [decimal]([math]::Round(($i * 11.317) % 100000, 2))
        $row.Enabled = ($i % 4) -ne 0
        $row.Notes = 'DataTable row {0}' -f $i
        $table.Rows.Add($row)
    }

    , $table
}

function New-BenchmarkDataSet {
    param([int] $Count)

    $dataSet = [Data.DataSet]::new('Report')
    $sales = New-BenchmarkDataTable -Count $Count
    $sales.TableName = 'Sales'
    $inventory = [Data.DataTable]::new('Inventory')
    $null = $inventory.Columns.Add('Sku', [string])
    $null = $inventory.Columns.Add('Quantity', [int])
    $null = $inventory.Columns.Add('Updated', [datetime])

    $inventoryCount = [math]::Max(1, [math]::Floor($Count / 4))
    for ($i = 1; $i -le $inventoryCount; $i++) {
        $row = $inventory.NewRow()
        $row.Sku = 'SKU-{0:000000}' -f $i
        $row.Quantity = $i % 500
        $row.Updated = ([datetime]'2024-01-01').AddHours($i)
        $inventory.Rows.Add($row)
    }

    $dataSet.Tables.Add($sales)
    $dataSet.Tables.Add($inventory)
    , $dataSet
}

function Get-BenchmarkData {
    param(
        [string] $Profile,
        [int] $Count
    )

    switch ($Profile) {
        'MixedObjects' {
            [pscustomobject]@{
                Data = @(New-BenchmarkRows -Count $Count)
                ColumnCount = 10
                WorksheetName = 'Data'
            }
            break
        }
        'WideObjects' {
            [pscustomobject]@{
                Data = @(New-WideBenchmarkRows -Count $Count)
                ColumnCount = 40
                WorksheetName = 'Data'
            }
            break
        }
        'DataTable' {
            [pscustomobject]@{
                Data = New-BenchmarkDataTable -Count $Count
                ColumnCount = 6
                WorksheetName = 'Data'
            }
            break
        }
        'DataSet' {
            [pscustomobject]@{
                Data = New-BenchmarkDataSet -Count $Count
                ColumnCount = 6
                WorksheetName = 'Sales'
            }
            break
        }
        default {
            throw "Unknown benchmark data profile '$Profile'."
        }
    }
}

function ConvertTo-ExcelColumnName {
    param([int] $ColumnNumber)

    $name = ''
    $value = $ColumnNumber
    while ($value -gt 0) {
        $value--
        $name = [char][int](65 + ($value % 26)) + $name
        $value = [int][math]::Floor($value / 26)
    }

    $name
}

function Get-RowCount {
    param([object] $Rows)

    if ($null -eq $Rows) { return 0 }
    if ($Rows -is [array]) { return $Rows.Count }
    if ($Rows -is [Data.DataTable]) { return $Rows.Rows.Count }
    if ($Rows -is [Data.DataSet]) {
        $count = 0
        foreach ($table in $Rows.Tables) {
            $count += $table.Rows.Count
        }
        return $count
    }

    return @($Rows).Count
}

function New-FollowUpScenario {
    param(
        [string] $Key,
        [string] $Name,
        [string[]] $Suites,
        [scriptblock] $Script
    )

    [pscustomobject]@{
        Key = $Key
        Name = $Name
        Suites = $Suites
        Script = $Script
    }
}

function New-ExportScenario {
    param(
        [string] $Key,
        [string] $Name,
        [string[]] $Suites,
        [string] $Engine,
        [string] $Profile,
        [string] $FileStem,
        [scriptblock] $Script,
        [object[]] $FollowUps = @()
    )

    [pscustomobject]@{
        Key = $Key
        Name = $Name
        Suites = $Suites
        Engine = $Engine
        Profile = $Profile
        FileStem = $FileStem
        Script = $Script
        FollowUps = @($FollowUps)
    }
}

function Get-ExcelBenchmarkScenarios {
    $basicSuites = @('Smoke', 'Standard', 'Large', 'Full', 'SuperLarge')
    $tableSuites = @('Smoke', 'Standard', 'Large', 'Full')
    $standardSuites = @('Standard', 'Large', 'Full')
    $scaleSuites = @('Standard', 'Large', 'Full', 'SuperLarge')
    $dataSetSuites = @('Large', 'Full')
    $reportSuites = @('Smoke', 'Standard', 'Large', 'Full')

    $defaultImport = New-FollowUpScenario -Key 'import-default-full' -Name 'Import full sheet from default export' -Suites $basicSuites -Script {
        param($Context)

        switch ($Context.Engine) {
            'PSWriteOffice' { Import-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName }
            'ImportExcel' { Import-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName }
            'ExcelFast' { Import-Workbook -Path $Context.Path -SheetName $Context.WorksheetName }
        }
    }

    $defaultRangeImport = New-FollowUpScenario -Key 'import-default-range' -Name 'Import A1 range from default export' -Suites $scaleSuites -Script {
        param($Context)

        switch ($Context.Engine) {
            'PSWriteOffice' { Import-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName -Range $Context.Range }
            'ImportExcel' { Import-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -StartRow 1 -EndRow ($Context.Rows + 1) -StartColumn 1 -EndColumn $Context.ColumnCount }
            'ExcelFast' { Import-Workbook -Path $Context.Path -SheetName $Context.WorksheetName -StartCell 'A1' -EndCell $Context.RangeEndCell }
        }
    }

    $tableImport = New-FollowUpScenario -Key 'import-table-full' -Name 'Import full sheet from table export' -Suites $tableSuites -Script {
        param($Context)

        switch ($Context.Engine) {
            'PSWriteOffice' { Import-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName }
            'ImportExcel' { Import-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName }
            'ExcelFast' { Import-Workbook -Path $Context.Path -SheetName $Context.WorksheetName }
        }
    }

    @(
        New-ExportScenario -Key 'objects-table' -Name 'Export objects as table' -Suites $tableSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-objects-table' -FollowUps @($tableImport) -Script {
            param($Context)
            $Context.Data | Export-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data
        }
        New-ExportScenario -Key 'objects-table' -Name 'Export objects as table' -Suites $tableSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-objects-table' -FollowUps @($tableImport) -Script {
            param($Context)
            $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -AutoFilter
        }
        New-ExportScenario -Key 'objects-default' -Name 'Export objects default' -Suites $basicSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-objects-default' -FollowUps @($defaultImport, $defaultRangeImport) -Script {
            param($Context)
            $Context.Data | Export-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName
        }
        New-ExportScenario -Key 'objects-default' -Name 'Export objects default' -Suites $basicSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-objects-default' -FollowUps @($defaultImport, $defaultRangeImport) -Script {
            param($Context)
            $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName
        }
        New-ExportScenario -Key 'objects-default' -Name 'Export objects default' -Suites $basicSuites -Engine 'ExcelFast' -Profile 'MixedObjects' -FileStem 'excelfast-objects-default' -FollowUps @($defaultImport, $defaultRangeImport) -Script {
            param($Context)
            Export-Workbook -Destination $Context.Path -InputObject $Context.Data -SheetName $Context.WorksheetName -Force
        }
        New-ExportScenario -Key 'objects-no-table' -Name 'Export objects no table' -Suites $scaleSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-objects-notable' -Script {
            param($Context)
            $Context.Data | Export-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName -NoTable
        }
        New-ExportScenario -Key 'objects-no-table' -Name 'Export objects no table' -Suites $scaleSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-objects-notable' -Script {
            param($Context)
            $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName
        }
        New-ExportScenario -Key 'objects-table-autofit' -Name 'Export objects table autofit' -Suites $standardSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-objects-table-autofit' -Script {
            param($Context)
            $Context.Data | Export-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -AutoFit
        }
        New-ExportScenario -Key 'objects-table-autofit' -Name 'Export objects table autofit' -Suites $standardSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-objects-table-autofit' -Script {
            param($Context)
            $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -AutoFilter -AutoSize
        }
        New-ExportScenario -Key 'wide-objects-default' -Name 'Export wide objects default' -Suites $scaleSuites -Engine 'PSWriteOffice' -Profile 'WideObjects' -FileStem 'pswriteoffice-wide-objects-default' -FollowUps @($defaultImport, $defaultRangeImport) -Script {
            param($Context)
            $Context.Data | Export-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName
        }
        New-ExportScenario -Key 'wide-objects-default' -Name 'Export wide objects default' -Suites $scaleSuites -Engine 'ImportExcel' -Profile 'WideObjects' -FileStem 'importexcel-wide-objects-default' -FollowUps @($defaultImport, $defaultRangeImport) -Script {
            param($Context)
            $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName
        }
        New-ExportScenario -Key 'wide-objects-default' -Name 'Export wide objects default' -Suites $scaleSuites -Engine 'ExcelFast' -Profile 'WideObjects' -FileStem 'excelfast-wide-objects-default' -FollowUps @($defaultImport, $defaultRangeImport) -Script {
            param($Context)
            Export-Workbook -Destination $Context.Path -InputObject $Context.Data -SheetName $Context.WorksheetName -Force
        }
        New-ExportScenario -Key 'datatable-default' -Name 'Export DataTable default' -Suites $scaleSuites -Engine 'PSWriteOffice' -Profile 'DataTable' -FileStem 'pswriteoffice-datatable-default' -FollowUps @($defaultImport, $defaultRangeImport) -Script {
            param($Context)
            Export-OfficeExcel -Path $Context.Path -InputObject $Context.Data -WorksheetName $Context.WorksheetName -TableName Data
        }
        New-ExportScenario -Key 'datatable-default' -Name 'Export DataTable default' -Suites $scaleSuites -Engine 'ImportExcel' -Profile 'DataTable' -FileStem 'importexcel-datatable-default' -FollowUps @($defaultImport, $defaultRangeImport) -Script {
            param($Context)
            $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName
        }
        New-ExportScenario -Key 'report-workbook' -Name 'Export report workbook with table chart pivot formatting' -Suites $reportSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-report-workbook' -FollowUps @($defaultImport, $defaultRangeImport) -Script {
            param($Context)

            $lastRow = $Context.Rows + 1
            $sourceRange = 'A1:{0}{1}' -f (ConvertTo-ExcelColumnName -ColumnNumber $Context.ColumnCount), $lastRow
            New-OfficeExcel -Path $Context.Path {
                Add-OfficeExcelSheet -Name $Context.WorksheetName -Content {
                    Add-OfficeExcelTable -Data $Context.Data -TableName Data -AutoFit
                    Set-OfficeExcelFreeze -TopRows 1
                    Add-OfficeExcelConditionalRule -Range "G2:G$lastRow" -Operator GreaterThan -Formula1 '750'
                    Add-OfficeExcelConditionalDataBar -Range "G2:G$lastRow" -Color '#70AD47'
                    Add-OfficeExcelConditionalColorScale -Range "I2:I$lastRow" -StartColor '#F4CCCC' -EndColor '#D9EAD3'
                    Add-OfficeExcelConditionalIconSet -Range "I2:I$lastRow"
                    Add-OfficeExcelValidationList -Range "D2:D$lastRow" -Values 'NA', 'EU', 'APAC', 'LATAM'
                    Set-OfficeExcelColumnStyleByHeader -Header Score -NumberFormat '#,##0.000'
                    Set-OfficeExcelColumnStyleByHeader -Header Created -NumberFormat 'yyyy-mm-dd hh:mm'
                    Add-OfficeExcelChart -TableName Data -Row 2 -Column 12 -Type ColumnClustered -Title 'Score by region'
                    Add-OfficeExcelPivotTable -SourceRange $sourceRange -DestinationCell 'L24' -Name 'SummaryPivot' -RowField Region -ColumnField Department -DataField Score, TicketCount -DataFunction Average, Sum -DataDisplayName 'Average Score', 'Tickets' -DataNumberFormat '#,##0.00', '#,##0'
                }
            } | Out-Null
        }
        New-ExportScenario -Key 'report-workbook' -Name 'Export report workbook with table chart pivot formatting' -Suites $reportSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-report-workbook' -FollowUps @($defaultImport, $defaultRangeImport) -Script {
            param($Context)

            $lastRow = $Context.Rows + 1
            $excel = $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -AutoFilter -AutoSize -FreezeTopRow -BoldTopRow -PassThru
            try {
                $worksheet = $excel.Workbook.Worksheets[$Context.WorksheetName]
                Add-ConditionalFormatting -Worksheet $worksheet -Address "G2:G$lastRow" -RuleType GreaterThan -ConditionValue 750 -BackgroundColor LightPink
                Add-ConditionalFormatting -Worksheet $worksheet -Address "G2:G$lastRow" -DataBarColor Green
                Add-ConditionalFormatting -Worksheet $worksheet -Address "I2:I$lastRow" -ThreeIconsSet TrafficLights1
                Add-ExcelDataValidationRule -Worksheet $worksheet -Range "D2:D$lastRow" -ValidationType List -ValueSet @('NA', 'EU', 'APAC', 'LATAM')
                $worksheet.Cells["G2:G$lastRow"].Style.Numberformat.Format = '#,##0.000'
                $worksheet.Cells["F2:F$lastRow"].Style.Numberformat.Format = 'yyyy-mm-dd hh:mm'
                Add-ExcelChart -Worksheet $worksheet -ChartType ColumnClustered -Title 'Score by region' -XRange "D2:D$lastRow" -YRange "G2:G$lastRow" -Row 2 -Column 12 -Width 640 -Height 360
                Add-PivotTable -ExcelPackage $excel -Address $worksheet.Cells['L24'] -SourceWorksheet $worksheet -SourceRange $worksheet.Tables[0].Address -PivotTableName SummaryPivot -PivotRows Region -PivotColumns Department -PivotData @{ Score = 'Average'; TicketCount = 'Sum' } -PivotNumberFormat '#,##0.00'
            } finally {
                Close-ExcelPackage -ExcelPackage $excel
            }
        }
        New-ExportScenario -Key 'dataset-worksheets' -Name 'Export DataSet worksheets' -Suites $dataSetSuites -Engine 'PSWriteOffice' -Profile 'DataSet' -FileStem 'pswriteoffice-dataset-worksheets' -FollowUps @($defaultImport) -Script {
            param($Context)
            Export-OfficeExcel -Path $Context.Path -InputObject $Context.Data
        }
    )
}

function Test-ScenarioFilter {
    param(
        [object] $ScenarioObject,
        [string[]] $Patterns
    )

    if (-not $Patterns -or $Patterns.Count -eq 0) {
        return $true
    }

    foreach ($pattern in $Patterns) {
        if ($ScenarioObject.Key -like $pattern -or $ScenarioObject.Name -like $pattern) {
            return $true
        }
    }

    return $false
}

function Get-SelectedFollowUps {
    param([object] $ScenarioObject)

    $followUps = @($ScenarioObject.FollowUps | Where-Object { $_.Suites -contains $Suite })
    if (-not $Scenario -or $Scenario.Count -eq 0) {
        return $followUps
    }

    if (Test-ScenarioFilter -ScenarioObject $ScenarioObject -Patterns $Scenario) {
        return $followUps
    }

    @($followUps | Where-Object { Test-ScenarioFilter -ScenarioObject $_ -Patterns $Scenario })
}

function Invoke-BenchmarkOperation {
    param(
        [object] $Context,
        [string] $ScenarioKey,
        [string] $ScenarioName,
        [scriptblock] $ScriptBlock
    )

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    $process = [Diagnostics.Process]::GetCurrentProcess()
    $beforeWorkingSet = $process.WorkingSet64
    $beforeManaged = [GC]::GetTotalMemory($false)
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    $resultCount = 0
    $status = 'Passed'
    $errorMessage = $null
    try {
        $result = & $ScriptBlock $Context
        $resultCount = Get-RowCount -Rows $result
    } catch {
        $status = 'Failed'
        $errorMessage = $_.Exception.Message
    }
    $stopwatch.Stop()
    $process.Refresh()

    [pscustomobject]@{
        TimestampUtc      = [datetime]::UtcNow.ToString('o')
        Suite             = $Suite
        Engine            = $Context.Engine
        ScenarioKey       = $ScenarioKey
        Scenario          = $ScenarioName
        Profile           = $Context.Profile
        Rows              = $Context.Rows
        Iteration         = $Context.Iteration
        Milliseconds      = [math]::Round($stopwatch.Elapsed.TotalMilliseconds, 3)
        ResultCount       = $resultCount
        FileBytes         = if (Test-Path $Context.Path) { (Get-Item $Context.Path).Length } else { 0 }
        WorkingSetDeltaMB = [math]::Round(($process.WorkingSet64 - $beforeWorkingSet) / 1MB, 3)
        ManagedDeltaMB    = [math]::Round(([GC]::GetTotalMemory($false) - $beforeManaged) / 1MB, 3)
        Status            = $status
        Error             = $errorMessage
    }
}

function Get-MedianValue {
    param(
        [object[]] $InputObject,
        [string] $PropertyName
    )

    $values = @(
        $InputObject |
            ForEach-Object {
                if ($_.PSObject.Properties[$PropertyName]) {
                    [double]$_.PSObject.Properties[$PropertyName].Value
                }
            } |
            Sort-Object
    )

    if ($values.Count -eq 0) {
        return 0
    }

    $values[[int][math]::Floor(($values.Count - 1) / 2)]
}

function Format-Ratio {
    param([double] $Value)

    if ($Value -le 0 -or [double]::IsNaN($Value) -or [double]::IsInfinity($Value)) {
        return $null
    }

    if ($Value -lt 10) {
        return ('{0:0.##} x' -f $Value)
    }

    ('{0:0.#} x' -f $Value)
}

function Get-ComparisonRating {
    param(
        [double] $Ratio,
        [int] $Rank
    )

    if ($Rank -eq 1) {
        return 'fastest'
    }
    if ($Ratio -le 1.15) {
        return 'competitive'
    }
    if ($Ratio -le 2) {
        return 'watch'
    }

    'behind'
}

function New-BenchmarkComparison {
    param([object[]] $Summary)

    $comparisonRows = [Collections.Generic.List[object]]::new()
    $groups = $Summary | Group-Object ScenarioKey, Scenario, Profile, Rows

    foreach ($group in $groups) {
        $passed = @(
            $group.Group |
                Where-Object { [int]$_.Passed -gt 0 -and [double]$_.MedianMs -gt 0 } |
                Sort-Object MedianMs
        )

        if ($passed.Count -eq 0) {
            continue
        }

        $fastest = $passed[0]
        $fastestMs = [double]$fastest.MedianMs
        $smallestFile = @($passed | Where-Object { [double]$_.MedianFileKB -gt 0 } | Sort-Object MedianFileKB | Select-Object -First 1)
        $pswriteOffice = @($passed | Where-Object Engine -eq 'PSWriteOffice' | Select-Object -First 1)
        $pswriteOfficeRank = 0
        $pswriteOfficeRatio = 0.0
        $pswriteOfficeText = 'not tested'
        $leadText = $null
        $rating = 'not tested'

        if ($pswriteOffice.Count -gt 0) {
            $pswriteOfficeRank = [array]::IndexOf($passed, $pswriteOffice[0]) + 1
            $pswriteOfficeRatio = [math]::Round(([double]$pswriteOffice[0].MedianMs) / $fastestMs, 4)
            if ($pswriteOfficeRank -eq 1) {
                if ($passed.Count -gt 1) {
                    $next = $passed[1]
                    $leadRatio = [math]::Round(([double]$next.MedianMs) / ([double]$pswriteOffice[0].MedianMs), 4)
                    $leadText = 'lead {0} vs {1}' -f (Format-Ratio -Value $leadRatio), $next.Engine
                    $pswriteOfficeText = '1 x (fastest, {0})' -f $leadText
                } else {
                    $pswriteOfficeText = '1 x (fastest)'
                }
            } else {
                $pswriteOfficeText = '{0} slower than {1}' -f (Format-Ratio -Value $pswriteOfficeRatio), $fastest.Engine
            }

            $rating = Get-ComparisonRating -Ratio $pswriteOfficeRatio -Rank $pswriteOfficeRank
        }

        $engineResults = @(
            $passed |
                ForEach-Object {
                    $timeRatio = [math]::Round(([double]$_.MedianMs) / $fastestMs, 4)
                    $fileRatio = if ($smallestFile.Count -gt 0 -and [double]$smallestFile[0].MedianFileKB -gt 0 -and [double]$_.MedianFileKB -gt 0) {
                        [math]::Round(([double]$_.MedianFileKB) / ([double]$smallestFile[0].MedianFileKB), 4)
                    } else {
                        $null
                    }

                    [pscustomobject]@{
                        Engine = $_.Engine
                        Rank = [array]::IndexOf($passed, $_) + 1
                        MedianMs = [double]$_.MedianMs
                        MinMs = [double]$_.MinMs
                        MaxMs = [double]$_.MaxMs
                        MedianFileKB = [double]$_.MedianFileKB
                        TimeVsFastest = $timeRatio
                        TimeVsFastestText = Format-Ratio -Value $timeRatio
                        FileVsSmallest = $fileRatio
                        FileVsSmallestText = if ($fileRatio) { Format-Ratio -Value $fileRatio } else { $null }
                    }
                }
        )

        $row = [ordered]@{
            ScenarioKey = $group.Group[0].ScenarioKey
            Scenario = $group.Group[0].Scenario
            Profile = $group.Group[0].Profile
            Rows = [int]$group.Group[0].Rows
            FastestEngine = $fastest.Engine
            FastestMs = $fastestMs
            PSWriteOfficeRank = $pswriteOfficeRank
            PSWriteOfficeVsFastest = $pswriteOfficeRatio
            PSWriteOfficeVsFastestText = $pswriteOfficeText
            LeadText = $leadText
            SmallestFileEngine = if ($smallestFile.Count -gt 0) { $smallestFile[0].Engine } else { $null }
            Rating = $rating
            Engines = $engineResults
        }

        foreach ($engineName in @('PSWriteOffice', 'ImportExcel', 'ExcelFast')) {
            $engineResult = @($engineResults | Where-Object Engine -eq $engineName | Select-Object -First 1)
            $prefix = $engineName -replace '[^A-Za-z0-9]', ''
            $row["${prefix}Ms"] = if ($engineResult.Count -gt 0) { $engineResult[0].MedianMs } else { $null }
            $row["${prefix}VsFastest"] = if ($engineResult.Count -gt 0) { $engineResult[0].TimeVsFastest } else { $null }
            $row["${prefix}FileKB"] = if ($engineResult.Count -gt 0) { $engineResult[0].MedianFileKB } else { $null }
        }

        $comparisonRows.Add([pscustomobject]$row)
    }

    $comparisonRows |
        Sort-Object ScenarioKey, Profile, Rows
}

if (-not $PSBoundParameters.ContainsKey('RowCount') -or -not $RowCount) {
    $RowCount = switch ($Suite) {
        'Smoke' { @(1000) }
        'Standard' { @(1000, 10000, 25000) }
        'Large' { @(25000, 100000, 250000) }
        'Full' { @(1000, 10000, 25000, 100000) }
        'SuperLarge' { @(250000, 500000, 1000000) }
    }
}
$RowCount = Resolve-RowCountList -Value $RowCount

if ($RepeatCount -le 0) {
    $RepeatCount = switch ($Suite) {
        'Smoke' { 1 }
        'Standard' { 3 }
        'Large' { 3 }
        'Full' { 5 }
        'SuperLarge' { 1 }
    }
}

$allScenarios = @(
    Get-ExcelBenchmarkScenarios |
        Where-Object { $_ -and $_.PSObject.Properties['Key'] -and $_.PSObject.Properties['Engine'] -and $_.PSObject.Properties['Script'] }
)
if ($ListScenarios.IsPresent) {
    $allScenarios |
        Sort-Object Key, Engine |
        ForEach-Object {
            [pscustomobject]@{
                Key = $_.Key
                Engine = $_.Engine
                Name = $_.Name
                Profile = $_.Profile
                Suites = ($_.Suites -join ', ')
                FollowUps = (@($_.FollowUps | Where-Object { $_ -and $_.PSObject.Properties['Key'] } | ForEach-Object { $_.Key }) -join ', ')
            }
        } |
        Format-Table -AutoSize
    return
}

$selectedScenarios = @(
    $allScenarios |
        Where-Object { $_.Suites -contains $Suite } |
        Where-Object { $Engine -contains $_.Engine } |
        Where-Object {
            if (-not $Scenario -or $Scenario.Count -eq 0) {
                return $true
            }

            if (Test-ScenarioFilter -ScenarioObject $_ -Patterns $Scenario) {
                return $true
            }

            $matchingFollowUps = @(Get-SelectedFollowUps -ScenarioObject $_)
            return $matchingFollowUps.Count -gt 0
        }
)

if ($selectedScenarios.Count -eq 0) {
    throw 'No benchmark scenarios matched the requested suite, engine, and scenario filters.'
}

$null = New-Item -ItemType Directory -Force -Path $moduleRoot, $workRoot

if ($Engine -contains 'ImportExcel') {
    Ensure-ImportExcel
}
if ($Engine -contains 'ExcelFast') {
    Ensure-ExcelFast
}

if ($Engine -contains 'PSWriteOffice') {
    $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
    $env:OfficeIMORoot = Join-Path $repoRoot '.missing-officeimo'
    Import-Module (Join-Path $repoRoot 'PSWriteOffice.psd1') -Force -ErrorAction Stop
}
if ($Engine -contains 'ImportExcel') {
    Import-Module ImportExcel -Force -ErrorAction Stop
}
if ($Engine -contains 'ExcelFast') {
    Import-Module ExcelFast -Force -ErrorAction Stop
}

$results = [Collections.Generic.List[object]]::new()
foreach ($rows in $RowCount) {
    $profileCache = @{}
    for ($iteration = 1; $iteration -le $RepeatCount; $iteration++) {
        foreach ($benchmarkScenario in $selectedScenarios) {
            if (-not $profileCache.ContainsKey($benchmarkScenario.Profile)) {
                $profileCache[$benchmarkScenario.Profile] = Get-BenchmarkData -Profile $benchmarkScenario.Profile -Count $rows
            }

            $profile = $profileCache[$benchmarkScenario.Profile]
            $path = Join-Path $workRoot ('{0}-{1}-{2}.xlsx' -f $benchmarkScenario.FileStem, $rows, $iteration)
            $rangeEndColumn = ConvertTo-ExcelColumnName -ColumnNumber $profile.ColumnCount
            $context = [pscustomobject]@{
                Engine = $benchmarkScenario.Engine
                Profile = $benchmarkScenario.Profile
                Data = $profile.Data
                ColumnCount = $profile.ColumnCount
                Rows = $rows
                Iteration = $iteration
                WorksheetName = $profile.WorksheetName
                Path = $path
                Range = 'A1:{0}{1}' -f $rangeEndColumn, ($rows + 1)
                RangeEndCell = '{0}{1}' -f $rangeEndColumn, ($rows + 1)
            }

            if (Test-Path $context.Path) {
                Remove-Item $context.Path -Force
            }

            $results.Add((Invoke-BenchmarkOperation -Context $context -ScenarioKey $benchmarkScenario.Key -ScenarioName $benchmarkScenario.Name -ScriptBlock $benchmarkScenario.Script))

            if ((-not $SkipFollowUps.IsPresent) -and (Test-Path $context.Path)) {
                foreach ($followUp in (Get-SelectedFollowUps -ScenarioObject $benchmarkScenario)) {
                    $results.Add((Invoke-BenchmarkOperation -Context $context -ScenarioKey $followUp.Key -ScenarioName $followUp.Name -ScriptBlock $followUp.Script))
                }
            }
        }
    }
}

$resultsPath = Join-Path $workRoot 'excel-performance-results.csv'
$summaryPath = Join-Path $workRoot 'excel-performance-summary.csv'
$comparisonCsvPath = Join-Path $workRoot 'excel-performance-comparison.csv'
$comparisonJsonPath = Join-Path $workRoot 'excel-performance-comparison.json'
$metadataPath = Join-Path $workRoot 'metadata.json'

$results | Export-Csv -NoTypeInformation -Path $resultsPath
$summary = $results |
    Group-Object Engine, ScenarioKey, Scenario, Profile, Rows |
    ForEach-Object {
        $passed = @($_.Group | Where-Object Status -eq 'Passed')
        $ordered = @($passed | Sort-Object Milliseconds)
        $median = if ($ordered.Count -eq 0) { 0 } else { $ordered[[int][math]::Floor(($ordered.Count - 1) / 2)].Milliseconds }
        [pscustomobject]@{
            Engine       = $_.Group[0].Engine
            ScenarioKey  = $_.Group[0].ScenarioKey
            Scenario     = $_.Group[0].Scenario
            Profile      = $_.Group[0].Profile
            Rows         = $_.Group[0].Rows
            Runs         = $_.Group.Count
            Passed       = $passed.Count
            MedianMs     = $median
            MinMs        = if ($passed.Count) { ($passed | Measure-Object Milliseconds -Minimum).Minimum } else { 0 }
            MaxMs        = if ($passed.Count) { ($passed | Measure-Object Milliseconds -Maximum).Maximum } else { 0 }
            MedianFileKB = if ($passed.Count) { [math]::Round((($passed | Sort-Object FileBytes)[[int][math]::Floor(($passed.Count - 1) / 2)].FileBytes) / 1KB, 1) } else { 0 }
            MedianWorkingSetDeltaMB = if ($passed.Count) { [math]::Round((Get-MedianValue -InputObject $passed -PropertyName WorkingSetDeltaMB), 3) } else { 0 }
            MedianManagedDeltaMB = if ($passed.Count) { [math]::Round((Get-MedianValue -InputObject $passed -PropertyName ManagedDeltaMB), 3) } else { 0 }
        }
    } |
    Sort-Object ScenarioKey, Profile, Rows, Engine

$summary | Export-Csv -NoTypeInformation -Path $summaryPath
$comparison = @(New-BenchmarkComparison -Summary $summary)
$comparison |
    Select-Object ScenarioKey, Scenario, Profile, Rows, FastestEngine, FastestMs, PSWriteOfficeMs, PSWriteOfficeRank, PSWriteOfficeVsFastest, PSWriteOfficeVsFastestText, LeadText, Rating, ImportExcelMs, ImportExcelVsFastest, ExcelFastMs, ExcelFastVsFastest, SmallestFileEngine, PSWriteOfficeFileKB, ImportExcelFileKB, ExcelFastFileKB |
    Export-Csv -NoTypeInformation -Path $comparisonCsvPath
$comparison | ConvertTo-Json -Depth 8 | Set-Content -Path $comparisonJsonPath -Encoding UTF8
$officeIMOExcelAssemblyPath = Join-Path $repoRoot 'Sources\PSWriteOffice\bin\Debug\net8.0\OfficeIMO.Excel.dll'
$officeIMOExcelAssemblyVersion = if (Test-Path $officeIMOExcelAssemblyPath) {
    [Reflection.AssemblyName]::GetAssemblyName($officeIMOExcelAssemblyPath).Version.ToString()
} else {
    $loadedOfficeIMOExcel = [AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GetName().Name -eq 'OfficeIMO.Excel' } |
        Select-Object -First 1
    if ($loadedOfficeIMOExcel) {
        $loadedOfficeIMOExcel.GetName().Version.ToString()
    } else {
        $null
    }
}

[pscustomobject]@{
    PowerShellVersion = $PSVersionTable.PSVersion.ToString()
    PSEdition = $PSEdition
    Suite = $Suite
    ImportExcel = if (Get-Module ImportExcel) { (Get-Module ImportExcel).Version.ToString() } else { $null }
    ExcelFast = if (Get-Module ExcelFast) { (Get-Module ExcelFast).Version.ToString() } else { $null }
    PSWriteOffice = if (Get-Module PSWriteOffice) { (Get-Module PSWriteOffice).Version.ToString() } else { $null }
    OfficeIMOExcelAssembly = $officeIMOExcelAssemblyVersion
    Engines = $Engine
    ScenarioFilter = $Scenario
    RowCount = $RowCount
    RepeatCount = $RepeatCount
    SkipFollowUps = $SkipFollowUps.IsPresent
    ScenarioCount = $selectedScenarios.Count
    ResultsPath = $resultsPath
    SummaryPath = $summaryPath
    ComparisonCsvPath = $comparisonCsvPath
    ComparisonJsonPath = $comparisonJsonPath
} | ConvertTo-Json -Depth 5 | Set-Content -Path $metadataPath -Encoding UTF8

Write-Host "Results: $resultsPath"
Write-Host "Summary: $summaryPath"
Write-Host "Comparison CSV: $comparisonCsvPath"
Write-Host "Comparison JSON: $comparisonJsonPath"
Write-Host "Metadata: $metadataPath"
$comparison |
    Select-Object ScenarioKey, Profile, Rows, FastestEngine, FastestMs, PSWriteOfficeMs, PSWriteOfficeVsFastestText, Rating |
    Format-Table -AutoSize
$summary | Format-Table -AutoSize
