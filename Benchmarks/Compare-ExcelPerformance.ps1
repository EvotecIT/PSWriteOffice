param(
    [ValidateSet('Smoke', 'Standard', 'Large', 'Full', 'SuperLarge')]
    [string] $Suite = 'Standard',

    [object[]] $RowCount,

    [int] $RepeatCount = 0,

    [string[]] $Scenario,

    [string[]] $Engine = @('PSWriteOffice', 'ImportExcel', 'ExcelFast', 'NativeCsv', 'CsvHelper'),

    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\Ignore\Benchmarks\ExcelPerformance'),

    [switch] $ListScenarios,

    [switch] $SkipFollowUps,

    [switch] $SkipWorkbookValidation,

    [switch] $SkipImportExcelInstall,

    [switch] $SkipExcelFastInstall,

    [switch] $SkipCsvHelperInstall,

    [string] $OfficeIMORoot,

    [ValidateSet('Debug', 'Release')]
    [string] $PSWriteOfficeConfiguration = 'Release',

    [switch] $SkipPSWriteOfficeBuild
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$invariantCulture = [Globalization.CultureInfo]::InvariantCulture
[Threading.Thread]::CurrentThread.CurrentCulture = $invariantCulture
[Threading.Thread]::CurrentThread.CurrentUICulture = $invariantCulture

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$moduleRoot = Join-Path $OutputDirectory 'Modules'
$packageRoot = Join-Path $OutputDirectory 'Packages'
$workRoot = Join-Path $OutputDirectory ('Run-{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $PID)
$validEngines = @('PSWriteOffice', 'ImportExcel', 'ExcelFast', 'NativeCsv', 'CsvHelper')

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

function Invoke-PSWriteOfficeBuild {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Debug', 'Release')]
        [string] $Configuration
    )

    $projectPath = Join-Path (Join-Path (Join-Path $repoRoot 'Sources') 'PSWriteOffice') 'PSWriteOffice.csproj'
    if (-not (Test-Path $projectPath)) {
        throw "PSWriteOffice project was not found at '$projectPath'."
    }

    Write-Host "Building PSWriteOffice ($Configuration) before benchmark import..."
    & dotnet build $projectPath -c $Configuration -v:minimal
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet build failed for PSWriteOffice ($Configuration)."
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

    try {
        Save-Module -Name ExcelFast -Path $moduleRoot -Repository PSGallery -AllowPrerelease -Force
    } catch {
        Write-Warning "ExcelFast could not be installed from PSGallery. Install it from https://github.com/JustinGrote/ExcelFast or place it on PSModulePath to include that lane. $($_.Exception.Message)"
        $script:Engine = @($script:Engine | Where-Object { $_ -ne 'ExcelFast' })
        return
    }

    Add-ModulePath -Path $moduleRoot
}

function Ensure-CsvHelper {
    if ([AppDomain]::CurrentDomain.GetAssemblies().GetName().Name -contains 'CsvHelper') {
        return
    }

    $packageName = 'CsvHelper'
    $packageVersion = '33.1.0'
    $globalPackageRoot = if ($env:NUGET_PACKAGES) {
        $env:NUGET_PACKAGES
    } else {
        Join-Path (Join-Path $HOME '.nuget') 'packages'
    }

    $packageFolder = Join-Path $globalPackageRoot ($packageName.ToLowerInvariant())
    $targetPackageFolder = Join-Path $packageFolder $packageVersion
    if (-not (Test-Path $targetPackageFolder)) {
        if ($SkipCsvHelperInstall.IsPresent) {
            throw "CsvHelper $packageVersion is not restored. Rerun without -SkipCsvHelperInstall to restore it into the NuGet package cache."
        }

        $cacheProjectRoot = Join-Path $packageRoot 'CsvHelperRestore'
        $null = New-Item -ItemType Directory -Force -Path $cacheProjectRoot
        $cacheProjectPath = Join-Path $cacheProjectRoot 'CsvHelperRestore.csproj'
        @"
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="$packageName" Version="$packageVersion" />
  </ItemGroup>
</Project>
"@ | Set-Content -Path $cacheProjectPath -Encoding UTF8

        dotnet restore $cacheProjectPath --nologo | Write-Verbose
        if ($LASTEXITCODE -ne 0) {
            throw "dotnet restore failed while restoring CsvHelper $packageVersion."
        }
    }

    $csvHelperAssemblyPath = Get-CsvHelperAssemblyPath -PackageFolder $packageFolder
    if (-not $csvHelperAssemblyPath) {
        throw "CsvHelper package was restored, but CsvHelper.dll could not be found under '$packageFolder'."
    }

    if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.Location -eq $csvHelperAssemblyPath })) {
        [Reflection.Assembly]::LoadFrom($csvHelperAssemblyPath) | Out-Null
    }

}

function Get-CsvHelperAssemblyPath {
    param([string] $PackageFolder)

    if (-not (Test-Path $PackageFolder)) {
        return $null
    }

    $frameworkPreference = @('net8.0', 'net7.0', 'net6.0', 'netstandard2.1', 'netstandard2.0', 'net47', 'net45')
    $versionFolders = Get-ChildItem -Path $PackageFolder -Directory |
        Sort-Object { [version]($_.Name -replace '-.*$', '') } -Descending

    foreach ($versionFolder in $versionFolders) {
        foreach ($framework in $frameworkPreference) {
            $candidate = Join-Path (Join-Path (Join-Path $versionFolder.FullName 'lib') $framework) 'CsvHelper.dll'
            if (Test-Path $candidate) {
                return $candidate
            }
        }
    }

    $fallback = Get-ChildItem -Path $PackageFolder -Recurse -Filter CsvHelper.dll |
        Sort-Object FullName |
        Select-Object -First 1
    if ($fallback) {
        return $fallback.FullName
    }

    $null
}

function Write-CsvHelperFile {
    param(
        [string] $Path,
        [object[]] $Data
    )

    $encoding = [Text.UTF8Encoding]::new($false)
    $writer = [IO.StreamWriter]::new($Path, $false, $encoding)
    $csv = [CsvHelper.CsvWriter]::new($writer, [Globalization.CultureInfo]::InvariantCulture)
    try {
        $headers = $null
        foreach ($item in $Data) {
            $row = ConvertTo-CsvHelperRow -InputObject $item -Headers $headers
            if ($null -eq $headers) {
                $headers = $row.Headers
                foreach ($header in $headers) {
                    $csv.WriteField([string]$header)
                }
                $csv.NextRecord()
            }

            foreach ($value in $row.Values) {
                $csv.WriteField($value)
            }
            $csv.NextRecord()
        }
    } finally {
        $csv.Dispose()
        $writer.Dispose()
    }
}

function Read-CsvHelperFile {
    param([string] $Path)

    $reader = [IO.StreamReader]::new($Path, [Text.Encoding]::UTF8, $true)
    $csv = [CsvHelper.CsvReader]::new($reader, [Globalization.CultureInfo]::InvariantCulture)
    try {
        $rows = [Collections.Generic.List[object]]::new()
        if (-not $csv.Read()) {
            return $rows
        }

        $csv.ReadHeader()
        $headers = @($csv.HeaderRecord)
        while ($csv.Read()) {
            $row = [ordered]@{}
            foreach ($header in $headers) {
                $row[$header] = $csv.GetField([string]$header)
            }

            $rows.Add([pscustomobject]$row)
        }

        $rows
    } finally {
        $csv.Dispose()
        $reader.Dispose()
    }
}

function ConvertTo-CsvHelperRow {
    param(
        [object] $InputObject,
        [string[]] $Headers
    )

    $psProperties = $InputObject.PSObject.Properties
    if ($null -eq $Headers) {
        $names = [Collections.Generic.List[string]]::new()
        $values = [Collections.Generic.List[object]]::new()
        foreach ($property in $psProperties) {
            if (-not $property.IsGettable -or [string]::IsNullOrWhiteSpace($property.Name)) {
                continue
            }

            $names.Add($property.Name)
            $values.Add($property.Value)
        }

        [pscustomobject]@{
            Headers = $names.ToArray()
            Values = $values.ToArray()
        }
        return
    }

    $rowValues = New-Object object[] $Headers.Count
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $property = $psProperties[$Headers[$i]]
        $rowValues[$i] = if ($property) { $property.Value } else { $null }
    }

    [pscustomobject]@{
        Headers = $Headers
        Values = $rowValues
    }
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

function Get-BenchmarkRegionGroups {
    param([object[]] $Rows)

    foreach ($region in @('NA', 'EU', 'APAC', 'LATAM')) {
        [pscustomobject]@{
            Name = $region
            TableName = 'Data_' + $region
            Data = @($Rows | Where-Object { $_.Region -eq $region })
        }
    }
}

function Get-BenchmarkSummaryRows {
    param([int] $Rows)

    @(
        [pscustomobject]@{ Metric = 'Rows'; Formula = 'COUNTA(Data!A2:A{0})' -f ($Rows + 1); NumberFormat = '#,##0' }
        [pscustomobject]@{ Metric = 'Average score'; Formula = 'AVERAGE(Data!G2:G{0})' -f ($Rows + 1); NumberFormat = '#,##0.00' }
        [pscustomobject]@{ Metric = 'Tickets'; Formula = 'SUM(Data!I2:I{0})' -f ($Rows + 1); NumberFormat = '#,##0' }
        [pscustomobject]@{ Metric = 'Enabled'; Formula = 'COUNTIF(Data!E2:E{0},TRUE)' -f ($Rows + 1); NumberFormat = '#,##0' }
    )
}

function Get-BenchmarkAppendSplit {
    param([object[]] $Rows)

    $initialCount = [math]::Max(1, [math]::Floor($Rows.Count * 0.8))
    [pscustomobject]@{
        Initial = @($Rows | Select-Object -First $initialCount)
        Append = @($Rows | Select-Object -Skip $initialCount)
    }
}

function Get-BenchmarkSmallSheetGroups {
    param(
        [object[]] $Rows,
        [int] $SheetCount = 20
    )

    $safeSheetCount = [math]::Max(1, [math]::Min($SheetCount, [math]::Max(1, $Rows.Count)))
    $rowsPerSheet = [math]::Max(1, [math]::Ceiling($Rows.Count / $safeSheetCount))
    for ($sheetIndex = 0; $sheetIndex -lt $safeSheetCount; $sheetIndex++) {
        $sheetRows = @($Rows | Select-Object -Skip ($sheetIndex * $rowsPerSheet) -First $rowsPerSheet)
        if ($sheetRows.Count -eq 0) {
            continue
        }

        [pscustomobject]@{
            Name = 'Sheet{0:00}' -f ($sheetIndex + 1)
            TableName = 'Data{0:00}' -f ($sheetIndex + 1)
            Data = $sheetRows
        }
    }
}

function Get-BenchmarkWorkbookMergeInput {
    param(
        [object[]] $Rows,
        [string] $BasePath
    )

    $firstCount = [math]::Max(1, [math]::Floor($Rows.Count / 2))
    [pscustomobject]@{
        SourceA = [IO.Path]::Combine([IO.Path]::GetDirectoryName($BasePath), ([IO.Path]::GetFileNameWithoutExtension($BasePath) + '.source-a.xlsx'))
        SourceB = [IO.Path]::Combine([IO.Path]::GetDirectoryName($BasePath), ([IO.Path]::GetFileNameWithoutExtension($BasePath) + '.source-b.xlsx'))
        RowsA = @($Rows | Select-Object -First $firstCount)
        RowsB = @($Rows | Select-Object -Skip $firstCount)
    }
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
        [scriptblock] $Script,
        [string[]] $Engines = @('PSWriteOffice', 'ImportExcel', 'ExcelFast')
    )

    [pscustomobject]@{
        Key = $Key
        Name = $Name
        Suites = $Suites
        Script = $Script
        Engines = $Engines
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
        [object[]] $FollowUps = @(),
        [scriptblock] $Setup,
        [string] $FileExtension = 'xlsx',
        [bool] $ValidateWorkbook = $true
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
        Setup = $Setup
        FileExtension = $FileExtension
        ValidateWorkbook = $ValidateWorkbook
    }
}

function Get-ExcelBenchmarkScenarios {
    $basicSuites = @('Smoke', 'Standard', 'Large', 'Full', 'SuperLarge')
    $tableSuites = @('Smoke', 'Standard', 'Large', 'Full')
    $standardSuites = @('Standard', 'Large', 'Full')
    $scaleSuites = @('Standard', 'Large', 'Full', 'SuperLarge')
    $dataSetSuites = @('Large', 'Full')
    $reportSuites = @('Smoke', 'Standard', 'Large', 'Full')
    $workflowSuites = @('Standard', 'Large', 'Full')
    $csvSuites = @('Smoke', 'Standard', 'Large', 'Full', 'SuperLarge')

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

    $noHeaderRangeImport = New-FollowUpScenario -Key 'read-no-header-range' -Name 'Read selected range without headers' -Suites $standardSuites -Script {
        param($Context)

        switch ($Context.Engine) {
            'PSWriteOffice' { Import-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName -Range $Context.Range -NoHeader }
            'ImportExcel' { Import-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -StartRow 1 -EndRow ($Context.Rows + 1) -StartColumn 1 -EndColumn $Context.ColumnCount -NoHeader }
            'ExcelFast' { Import-Workbook -Path $Context.Path -SheetName $Context.WorksheetName -StartCell 'A1' -EndCell $Context.RangeEndCell -NoHeaders }
        }
    }

    $usedRangeAsDataTable = New-FollowUpScenario -Key 'read-used-range-datatable' -Name 'Read used range as DataTable' -Suites $standardSuites -Engines @('PSWriteOffice') -Script {
        param($Context)

        Get-OfficeExcelUsedRange -Path $Context.Path -Sheet $Context.WorksheetName -AsDataTable
    }

    $tableMetadataRead = New-FollowUpScenario -Key 'read-table-metadata' -Name 'Read workbook table metadata' -Suites $standardSuites -Engines @('PSWriteOffice') -Script {
        param($Context)

        Get-OfficeExcelTable -Path $Context.Path -Sheet $Context.WorksheetName
    }

    $namedRangeRead = New-FollowUpScenario -Key 'read-named-range-metadata' -Name 'Read workbook named range metadata' -Suites $standardSuites -Engines @('PSWriteOffice') -Script {
        param($Context)

        Get-OfficeExcelNamedRange -Path $Context.Path -Sheet $Context.WorksheetName
    }

    $csvImport = New-FollowUpScenario -Key 'csv-read' -Name 'Read CSV file' -Suites $csvSuites -Engines @('PSWriteOffice', 'NativeCsv', 'CsvHelper') -Script {
        param($Context)

        switch ($Context.Engine) {
            'PSWriteOffice' { Import-OfficeCsv -Path $Context.Path }
            'NativeCsv' { Import-Csv -Path $Context.Path }
            'CsvHelper' { Read-CsvHelperFile -Path $Context.Path }
        }
    }

    @(
        New-ExportScenario -Key 'csv-write' -Name 'Write CSV file' -Suites $csvSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-csv-write' -FileExtension 'csv' -ValidateWorkbook $false -FollowUps @($csvImport) -Script {
            param($Context)
            $Context.Data | Export-OfficeCsv -Path $Context.Path
        }
        New-ExportScenario -Key 'csv-write' -Name 'Write CSV file' -Suites $csvSuites -Engine 'NativeCsv' -Profile 'MixedObjects' -FileStem 'nativecsv-csv-write' -FileExtension 'csv' -ValidateWorkbook $false -FollowUps @($csvImport) -Script {
            param($Context)
            $Context.Data | Export-Csv -Path $Context.Path -NoTypeInformation -Encoding utf8 -UseQuotes AsNeeded
        }
        New-ExportScenario -Key 'csv-write' -Name 'Write CSV file' -Suites $csvSuites -Engine 'CsvHelper' -Profile 'MixedObjects' -FileStem 'csvhelper-csv-write' -FileExtension 'csv' -ValidateWorkbook $false -FollowUps @($csvImport) -Script {
            param($Context)
            Write-CsvHelperFile -Path $Context.Path -Data $Context.Data
        }
        New-ExportScenario -Key 'csv-read-source' -Name 'Read external CSV source' -Suites $csvSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-csv-read-source' -FileExtension 'csv' -ValidateWorkbook $false -Setup {
            param($Context)
            $Context.Data | Export-Csv -Path $Context.SourcePath -NoTypeInformation -Encoding utf8 -UseQuotes AsNeeded
        } -Script {
            param($Context)
            Import-OfficeCsv -Path $Context.SourcePath
        }
        New-ExportScenario -Key 'csv-read-source' -Name 'Read external CSV source' -Suites $csvSuites -Engine 'NativeCsv' -Profile 'MixedObjects' -FileStem 'nativecsv-csv-read-source' -FileExtension 'csv' -ValidateWorkbook $false -Setup {
            param($Context)
            $Context.Data | Export-Csv -Path $Context.SourcePath -NoTypeInformation -Encoding utf8 -UseQuotes AsNeeded
        } -Script {
            param($Context)
            Import-Csv -Path $Context.SourcePath
        }
        New-ExportScenario -Key 'csv-read-source' -Name 'Read external CSV source' -Suites $csvSuites -Engine 'CsvHelper' -Profile 'MixedObjects' -FileStem 'csvhelper-csv-read-source' -FileExtension 'csv' -ValidateWorkbook $false -Setup {
            param($Context)
            $Context.Data | Export-Csv -Path $Context.SourcePath -NoTypeInformation -Encoding utf8 -UseQuotes AsNeeded
        } -Script {
            param($Context)
            Read-CsvHelperFile -Path $Context.SourcePath
        }
        New-ExportScenario -Key 'csv-to-excel' -Name 'Create workbook from CSV source' -Suites $workflowSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-csv-to-excel' -Setup {
            param($Context)
            $Context.Data | Export-Csv -Path $Context.SourcePath -NoTypeInformation -Encoding utf8
        } -Script {
            param($Context)
            Import-OfficeExcelDelimitedText -Path $Context.Path -SourcePath $Context.SourcePath -SheetName $Context.WorksheetName
        }
        New-ExportScenario -Key 'csv-to-excel' -Name 'Create workbook from CSV source' -Suites $workflowSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-csv-to-excel' -Setup {
            param($Context)
            $Context.Data | Export-Csv -Path $Context.SourcePath -NoTypeInformation -Encoding utf8
        } -Script {
            param($Context)
            Import-Csv -Path $Context.SourcePath | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName
        }
        New-ExportScenario -Key 'objects-table' -Name 'Export objects as table' -Suites $tableSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-objects-table' -FollowUps @($tableImport) -Script {
            param($Context)
            $Context.Data | Export-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data
        }
        New-ExportScenario -Key 'objects-table' -Name 'Export objects as table' -Suites $tableSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-objects-table' -FollowUps @($tableImport) -Script {
            param($Context)
            $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -AutoFilter
        }
        New-ExportScenario -Key 'objects-default' -Name 'Export objects default' -Suites $basicSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-objects-default' -FollowUps @($defaultImport, $defaultRangeImport, $noHeaderRangeImport, $usedRangeAsDataTable, $tableMetadataRead) -Script {
            param($Context)
            $Context.Data | Export-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName
        }
        New-ExportScenario -Key 'objects-default' -Name 'Export objects default' -Suites $basicSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-objects-default' -FollowUps @($defaultImport, $defaultRangeImport, $noHeaderRangeImport, $usedRangeAsDataTable, $tableMetadataRead) -Script {
            param($Context)
            $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName
        }
        New-ExportScenario -Key 'objects-default' -Name 'Export objects default' -Suites $basicSuites -Engine 'ExcelFast' -Profile 'MixedObjects' -FileStem 'excelfast-objects-default' -FollowUps @($defaultImport, $defaultRangeImport, $noHeaderRangeImport, $usedRangeAsDataTable, $tableMetadataRead) -Script {
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
        New-ExportScenario -Key 'objects-title-freeze' -Name 'Export objects with title, offset header, and frozen top row' -Suites $workflowSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-objects-title-freeze' -Script {
            param($Context)
            $Context.Data | Export-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -Title 'Operational export' -StartRow 3 -FreezeTopRow -BoldTopRow
        }
        New-ExportScenario -Key 'objects-title-freeze' -Name 'Export objects with title, offset header, and frozen top row' -Suites $workflowSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-objects-title-freeze' -Script {
            param($Context)
            $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -AutoFilter -Title 'Operational export' -StartRow 3 -FreezeTopRow -BoldTopRow
        }
        New-ExportScenario -Key 'wide-objects-default' -Name 'Export wide objects default' -Suites $scaleSuites -Engine 'PSWriteOffice' -Profile 'WideObjects' -FileStem 'pswriteoffice-wide-objects-default' -FollowUps @($defaultImport, $defaultRangeImport, $noHeaderRangeImport, $usedRangeAsDataTable, $tableMetadataRead) -Script {
            param($Context)
            $Context.Data | Export-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName
        }
        New-ExportScenario -Key 'wide-objects-default' -Name 'Export wide objects default' -Suites $scaleSuites -Engine 'ImportExcel' -Profile 'WideObjects' -FileStem 'importexcel-wide-objects-default' -FollowUps @($defaultImport, $defaultRangeImport, $noHeaderRangeImport, $usedRangeAsDataTable, $tableMetadataRead) -Script {
            param($Context)
            $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName
        }
        New-ExportScenario -Key 'wide-objects-default' -Name 'Export wide objects default' -Suites $scaleSuites -Engine 'ExcelFast' -Profile 'WideObjects' -FileStem 'excelfast-wide-objects-default' -FollowUps @($defaultImport, $defaultRangeImport, $noHeaderRangeImport, $usedRangeAsDataTable, $tableMetadataRead) -Script {
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
        New-ExportScenario -Key 'multi-sheet-regions' -Name 'Export regional workbook with one table per sheet' -Suites $workflowSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-multi-sheet-regions' -Script {
            param($Context)

            New-OfficeExcel -Path $Context.Path {
                foreach ($group in (Get-BenchmarkRegionGroups -Rows $Context.Data)) {
                    Add-OfficeExcelSheet -Name $group.Name -Content {
                        Add-OfficeExcelTable -Data $group.Data -TableName $group.TableName
                        Set-OfficeExcelFreeze -TopRows 1
                    }
                }
            } | Out-Null
        }
        New-ExportScenario -Key 'multi-sheet-regions' -Name 'Export regional workbook with one table per sheet' -Suites $workflowSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-multi-sheet-regions' -Script {
            param($Context)

            $excel = $null
            try {
                foreach ($group in (Get-BenchmarkRegionGroups -Rows $Context.Data)) {
                    if ($excel) {
                        $excel = $group.Data | Export-Excel -ExcelPackage $excel -WorksheetName $group.Name -TableName $group.TableName -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
                    } else {
                        $excel = $group.Data | Export-Excel -Path $Context.Path -WorksheetName $group.Name -TableName $group.TableName -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
                    }
                }
            } finally {
                if ($excel) {
                    Close-ExcelPackage -ExcelPackage $excel
                }
            }
        }
        New-ExportScenario -Key 'summary-formulas' -Name 'Export data workbook with summary formulas' -Suites $workflowSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-summary-formulas' -Script {
            param($Context)

            $summaryRows = Get-BenchmarkSummaryRows -Rows $Context.Rows
            New-OfficeExcel -Path $Context.Path {
                Add-OfficeExcelSheet -Name $Context.WorksheetName -Content {
                    Add-OfficeExcelTable -Data $Context.Data -TableName Data
                    Set-OfficeExcelFreeze -TopRows 1
                }
                Add-OfficeExcelSheet -Name 'Summary' -Content {
                    Set-OfficeExcelCell -Address 'A1' -Value 'Metric'
                    Set-OfficeExcelCell -Address 'B1' -Value 'Value'
                    $row = 2
                    foreach ($summaryRow in $summaryRows) {
                        Set-OfficeExcelCell -Row $row -Column 1 -Value $summaryRow.Metric
                        Set-OfficeExcelFormula -Address ('B{0}' -f $row) -Formula $summaryRow.Formula
                        Set-OfficeExcelCell -Row $row -Column 2 -NumberFormat $summaryRow.NumberFormat
                        $row++
                    }
                }
            } | Out-Null
        }
        New-ExportScenario -Key 'summary-formulas' -Name 'Export data workbook with summary formulas' -Suites $workflowSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-summary-formulas' -Script {
            param($Context)

            $summaryRows = Get-BenchmarkSummaryRows -Rows $Context.Rows
            $excel = $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
            try {
                $summary = $excel.Workbook.Worksheets.Add('Summary')
                $summary.Cells['A1'].Value = 'Metric'
                $summary.Cells['B1'].Value = 'Value'
                $row = 2
                foreach ($summaryRow in $summaryRows) {
                    $summary.Cells[$row, 1].Value = $summaryRow.Metric
                    $summary.Cells[$row, 2].Formula = $summaryRow.Formula
                    $summary.Cells[$row, 2].Style.Numberformat.Format = $summaryRow.NumberFormat
                    $row++
                }
            } finally {
                Close-ExcelPackage -ExcelPackage $excel
            }
        }
        New-ExportScenario -Key 'append-existing-table' -Name 'Append rows to an existing workbook table' -Suites $workflowSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-append-existing-table' -FollowUps @($defaultImport, $tableMetadataRead) -Script {
            param($Context)

            $split = Get-BenchmarkAppendSplit -Rows $Context.Data
            $split.Initial | Export-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data
            if ($split.Append.Count -gt 0) {
                Add-OfficeExcelTableRow -InputPath $Context.Path -Sheet $Context.WorksheetName -TableName Data -InputObject $split.Append | Out-Null
            }
        }
        New-ExportScenario -Key 'append-existing-table' -Name 'Append rows to an existing workbook table' -Suites $workflowSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-append-existing-table' -FollowUps @($defaultImport, $tableMetadataRead) -Script {
            param($Context)

            $split = Get-BenchmarkAppendSplit -Rows $Context.Data
            $split.Initial | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -AutoFilter
            if ($split.Append.Count -gt 0) {
                $split.Append | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -Append
            }
        }
        New-ExportScenario -Key 'update-existing-workbook' -Name 'Update cells and formulas in an existing workbook' -Suites $workflowSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-update-existing-workbook' -FollowUps @($defaultImport, $defaultRangeImport) -Script {
            param($Context)

            $Context.Data | Export-OfficeExcel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data
            Edit-OfficeExcelRow -InputPath $Context.Path -Sheet $Context.WorksheetName -ScriptBlock {
                param($row)

                $row.Set('TicketCount', ([int]$row.Get[int]('TicketCount') + 1))
                if ($row.Get[bool]('IsEnabled')) {
                    $row.Set('Notes', 'Reviewed')
                }
            } | Out-Null

            $document = Get-OfficeExcel -Path $Context.Path
            try {
                $sheet = $document.Sheets | Where-Object { $_.Name -eq $Context.WorksheetName } | Select-Object -First 1
                $formulaColumn = $Context.ColumnCount + 1
                $sheet.Cell(1, $formulaColumn, 'ScoreDouble', $null, $null)
                $sheet.Cell(2, $formulaColumn, $null, 'G2*2', '#,##0.00')
                $document | Save-OfficeExcel
            } finally {
                $document | Close-OfficeExcel
            }
        }
        New-ExportScenario -Key 'update-existing-workbook' -Name 'Update cells and formulas in an existing workbook' -Suites $workflowSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-update-existing-workbook' -FollowUps @($defaultImport, $defaultRangeImport) -Script {
            param($Context)

            $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -AutoFilter
            $excel = Open-ExcelPackage -Path $Context.Path
            try {
                $sheet = $excel.Workbook.Worksheets[$Context.WorksheetName]
                for ($row = 2; $row -le ($Context.Rows + 1); $row++) {
                    $sheet.Cells[$row, 9].Value = [int]$sheet.Cells[$row, 9].Value + 1
                    if ([bool]$sheet.Cells[$row, 5].Value) {
                        $sheet.Cells[$row, 10].Value = 'Reviewed'
                    }
                }

                $formulaColumn = $Context.ColumnCount + 1
                $sheet.Cells[1, $formulaColumn].Value = 'ScoreDouble'
                $sheet.Cells[2, $formulaColumn].Formula = 'G2*2'
                $sheet.Cells[2, $formulaColumn].Style.Numberformat.Format = '#,##0.00'
            } finally {
                Close-ExcelPackage -ExcelPackage $excel
            }
        }
        New-ExportScenario -Key 'many-small-sheets' -Name 'Export many small worksheets' -Suites $workflowSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-many-small-sheets' -Script {
            param($Context)

            New-OfficeExcel -Path $Context.Path {
                foreach ($group in (Get-BenchmarkSmallSheetGroups -Rows $Context.Data -SheetCount 20)) {
                    Add-OfficeExcelSheet -Name $group.Name -Content {
                        Add-OfficeExcelTable -Data $group.Data -TableName $group.TableName
                        Set-OfficeExcelFreeze -TopRows 1
                    }
                }
            } | Out-Null
        }
        New-ExportScenario -Key 'many-small-sheets' -Name 'Export many small worksheets' -Suites $workflowSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-many-small-sheets' -Script {
            param($Context)

            $excel = $null
            try {
                foreach ($group in (Get-BenchmarkSmallSheetGroups -Rows $Context.Data -SheetCount 20)) {
                    if ($excel) {
                        $excel = $group.Data | Export-Excel -ExcelPackage $excel -WorksheetName $group.Name -TableName $group.TableName -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
                    } else {
                        $excel = $group.Data | Export-Excel -Path $Context.Path -WorksheetName $group.Name -TableName $group.TableName -AutoFilter -FreezeTopRow -BoldTopRow -PassThru
                    }
                }
            } finally {
                if ($excel) {
                    Close-ExcelPackage -ExcelPackage $excel
                }
            }
        }
        New-ExportScenario -Key 'workbook-package-merge' -Name 'Merge workbook sheets with package copy' -Suites $workflowSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-workbook-package-merge' -Script {
            param($Context)

            $input = Get-BenchmarkWorkbookMergeInput -Rows $Context.Data -BasePath $Context.Path
            $input.RowsA | Export-OfficeExcel -Path $input.SourceA -WorksheetName 'Data' -TableName 'DataA'
            $input.RowsB | Export-OfficeExcel -Path $input.SourceB -WorksheetName 'Data' -TableName 'DataB'
            Join-OfficeExcelWorkbook -Path $Context.Path -SourcePath @($input.SourceA, $input.SourceB) -CopyMode Package -SheetNamePrefix 'Merged' | Out-Null
        }
        New-ExportScenario -Key 'workbook-package-merge' -Name 'Merge workbook sheets with package copy' -Suites $workflowSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-workbook-package-merge' -Script {
            param($Context)

            $input = Get-BenchmarkWorkbookMergeInput -Rows $Context.Data -BasePath $Context.Path
            $input.RowsA | Export-Excel -Path $input.SourceA -WorksheetName 'Data' -TableName 'DataA' -AutoFilter
            $input.RowsB | Export-Excel -Path $input.SourceB -WorksheetName 'Data' -TableName 'DataB' -AutoFilter

            $targetPackage = [OfficeOpenXml.ExcelPackage]::new([IO.FileInfo]$Context.Path)
            try {
                foreach ($item in @(
                    [pscustomobject]@{ Path = $input.SourceA; Name = 'MergedDataA' }
                    [pscustomobject]@{ Path = $input.SourceB; Name = 'MergedDataB' }
                )) {
                    $sourcePackage = [OfficeOpenXml.ExcelPackage]::new([IO.FileInfo]$item.Path)
                    try {
                        $sourceSheet = $sourcePackage.Workbook.Worksheets['Data']
                        if ($null -eq $sourceSheet) {
                            throw "Source worksheet 'Data' was not found in '$($item.Path)'."
                        }

                        $null = $targetPackage.Workbook.Worksheets.Add($item.Name, $sourceSheet)
                    } finally {
                        $sourcePackage.Dispose()
                    }
                }

                $targetPackage.Save()
            } finally {
                $targetPackage.Dispose()
            }
        }
        New-ExportScenario -Key 'named-range-workbook' -Name 'Export workbook with named data range' -Suites $workflowSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-named-range-workbook' -FollowUps @($defaultImport, $namedRangeRead) -Script {
            param($Context)

            $sourceRange = 'A1:{0}{1}' -f (ConvertTo-ExcelColumnName -ColumnNumber $Context.ColumnCount), ($Context.Rows + 1)
            New-OfficeExcel -Path $Context.Path {
                Add-OfficeExcelSheet -Name $Context.WorksheetName -Content {
                    Add-OfficeExcelTable -Data $Context.Data -TableName Data
                    Set-OfficeExcelNamedRange -Name SalesData -Range $sourceRange
                }
            } | Out-Null
        }
        New-ExportScenario -Key 'named-range-workbook' -Name 'Export workbook with named data range' -Suites $workflowSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-named-range-workbook' -FollowUps @($defaultImport, $namedRangeRead) -Script {
            param($Context)

            $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -AutoFilter -RangeName SalesData
        }
        New-ExportScenario -Key 'chart-only-workbook' -Name 'Export workbook with table and chart' -Suites $workflowSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-chart-only-workbook' -Script {
            param($Context)

            New-OfficeExcel -Path $Context.Path {
                Add-OfficeExcelSheet -Name $Context.WorksheetName -Content {
                    Add-OfficeExcelTable -Data $Context.Data -TableName Data
                    Add-OfficeExcelChart -TableName Data -Row 2 -Column 12 -Type ColumnClustered -Title 'Score by region'
                }
            } | Out-Null
        }
        New-ExportScenario -Key 'chart-only-workbook' -Name 'Export workbook with table and chart' -Suites $workflowSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-chart-only-workbook' -Script {
            param($Context)

            $lastRow = $Context.Rows + 1
            $excel = $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -AutoFilter -PassThru
            try {
                $worksheet = $excel.Workbook.Worksheets[$Context.WorksheetName]
                Add-ExcelChart -Worksheet $worksheet -ChartType ColumnClustered -Title 'Score by region' -XRange "D2:D$lastRow" -YRange "G2:G$lastRow" -Row 2 -Column 12 -Width 640 -Height 360
            } finally {
                Close-ExcelPackage -ExcelPackage $excel
            }
        }
        New-ExportScenario -Key 'pivot-only-workbook' -Name 'Export workbook with table and pivot' -Suites $workflowSuites -Engine 'PSWriteOffice' -Profile 'MixedObjects' -FileStem 'pswriteoffice-pivot-only-workbook' -Script {
            param($Context)

            $lastRow = $Context.Rows + 1
            $sourceRange = 'A1:{0}{1}' -f (ConvertTo-ExcelColumnName -ColumnNumber $Context.ColumnCount), $lastRow
            New-OfficeExcel -Path $Context.Path {
                Add-OfficeExcelSheet -Name $Context.WorksheetName -Content {
                    Add-OfficeExcelTable -Data $Context.Data -TableName Data
                    Add-OfficeExcelPivotTable -SourceRange $sourceRange -DestinationCell 'L4' -Name 'SummaryPivot' -RowField Region -ColumnField Department -DataField Score, TicketCount -DataFunction Average, Sum -DataDisplayName 'Average Score', 'Tickets' -DataNumberFormat '#,##0.00', '#,##0'
                }
            } | Out-Null
        }
        New-ExportScenario -Key 'pivot-only-workbook' -Name 'Export workbook with table and pivot' -Suites $workflowSuites -Engine 'ImportExcel' -Profile 'MixedObjects' -FileStem 'importexcel-pivot-only-workbook' -Script {
            param($Context)

            $excel = $Context.Data | Export-Excel -Path $Context.Path -WorksheetName $Context.WorksheetName -TableName Data -AutoFilter -PassThru
            try {
                $worksheet = $excel.Workbook.Worksheets[$Context.WorksheetName]
                Add-PivotTable -ExcelPackage $excel -Address $worksheet.Cells['L4'] -SourceWorksheet $worksheet -SourceRange $worksheet.Tables[0].Address -PivotTableName SummaryPivot -PivotRows Region -PivotColumns Department -PivotData @{ Score = 'Average'; TicketCount = 'Sum' } -PivotNumberFormat '#,##0.00'
            } finally {
                Close-ExcelPackage -ExcelPackage $excel
            }
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
                Add-ConditionalFormatting -Worksheet $worksheet -Address "I2:I$lastRow" -RuleType ThreeColorScale
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

    $followUps = @(
        $ScenarioObject.FollowUps |
            Where-Object { $_.Suites -contains $Suite } |
            Where-Object { -not $_.PSObject.Properties['Engines'] -or $_.Engines -contains $ScenarioObject.Engine }
    )
    if (-not $Scenario -or $Scenario.Count -eq 0) {
        return $followUps
    }

    if (Test-ScenarioFilter -ScenarioObject $ScenarioObject -Patterns $Scenario) {
        return $followUps
    }

    @($followUps | Where-Object { Test-ScenarioFilter -ScenarioObject $_ -Patterns $Scenario })
}

function Get-BenchmarkLoadedType {
    param([string] $FullName)

    foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
        $type = $assembly.GetType($FullName, $false, $false)
        if ($type) {
            return $type
        }
    }

    $null
}

function Test-BenchmarkWorkbook {
    param([string] $Path)

    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    $openStatus = 'Skipped'
    $openXmlStatus = 'Skipped'
    $errorMessage = $null

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path $Path)) {
        $stopwatch.Stop()
        return [pscustomobject]@{
            Status = 'Skipped'
            OpenStatus = $openStatus
            OpenXmlStatus = $openXmlStatus
            Milliseconds = [math]::Round($stopwatch.Elapsed.TotalMilliseconds, 3)
            Error = 'Workbook file was not created.'
        }
    }

    try {
        $document = Get-OfficeExcel -Path $Path -ReadOnly
        if ($document) {
            Close-OfficeExcel -Document $document
        }
        $openStatus = 'Passed'
    } catch {
        $openStatus = 'Failed'
        $errorMessage = $_.Exception.Message
    }

    if ($openStatus -eq 'Passed') {
        $spreadsheetDocumentType = Get-BenchmarkLoadedType -FullName 'DocumentFormat.OpenXml.Packaging.SpreadsheetDocument'
        $openXmlValidatorType = Get-BenchmarkLoadedType -FullName 'DocumentFormat.OpenXml.Validation.OpenXmlValidator'

        if ($spreadsheetDocumentType -and $openXmlValidatorType) {
            $spreadsheetDocument = $null
            try {
                $openMethod = $spreadsheetDocumentType.GetMethod('Open', [type[]]@([string], [bool]))
                $spreadsheetDocument = $openMethod.Invoke($null, @($Path, $false))
                $validator = [Activator]::CreateInstance($openXmlValidatorType)
                $validationErrors = @($validator.Validate($spreadsheetDocument) | Select-Object -First 1)
                if ($validationErrors.Count -gt 0) {
                    $openXmlStatus = 'Failed'
                    $errorMessage = $validationErrors[0].Description
                } else {
                    $openXmlStatus = 'Passed'
                }
            } catch {
                $openXmlStatus = 'Failed'
                $errorMessage = $_.Exception.Message
            } finally {
                if ($spreadsheetDocument) {
                    $spreadsheetDocument.Dispose()
                }
            }
        }
    }

    $stopwatch.Stop()
    $status = if ($openStatus -eq 'Failed' -or $openXmlStatus -eq 'Failed') {
        'Failed'
    } elseif ($openStatus -eq 'Passed') {
        'Passed'
    } else {
        'Skipped'
    }

    [pscustomobject]@{
        Status = $status
        OpenStatus = $openStatus
        OpenXmlStatus = $openXmlStatus
        Milliseconds = [math]::Round($stopwatch.Elapsed.TotalMilliseconds, 3)
        Error = $errorMessage
    }
}

function Invoke-BenchmarkOperation {
    param(
        [object] $Context,
        [string] $ScenarioKey,
        [string] $ScenarioName,
        [scriptblock] $ScriptBlock,
        [bool] $ValidateWorkbook = $false
    )

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    $process = [Diagnostics.Process]::GetCurrentProcess()
    $process.Refresh()
    $beforeWorkingSet = $process.WorkingSet64
    $beforePeakWorkingSet = $process.PeakWorkingSet64
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
    $afterWorkingSet = $process.WorkingSet64
    $afterPeakWorkingSet = $process.PeakWorkingSet64
    $artifactPath = if (Test-Path -LiteralPath $Context.Path) {
        $Context.Path
    } elseif ($Context.PSObject.Properties['SourcePath'] -and (Test-Path -LiteralPath $Context.SourcePath)) {
        $Context.SourcePath
    } else {
        $null
    }
    $workbookValidation = if ($ValidateWorkbook -and $status -eq 'Passed') {
        Test-BenchmarkWorkbook -Path $Context.Path
    } else {
        [pscustomobject]@{
            Status = 'Skipped'
            OpenStatus = 'Skipped'
            OpenXmlStatus = 'Skipped'
            Milliseconds = 0
            Error = $null
        }
    }

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
        FileBytes         = if ($artifactPath) { [long](Get-Item -LiteralPath $artifactPath).Length } else { 0L }
        WorkingSetBeforeMB = [math]::Round($beforeWorkingSet / 1MB, 3)
        WorkingSetAfterMB = [math]::Round($afterWorkingSet / 1MB, 3)
        WorkingSetDeltaMB = [math]::Round(($afterWorkingSet - $beforeWorkingSet) / 1MB, 3)
        PeakWorkingSetMB = [math]::Round($afterPeakWorkingSet / 1MB, 3)
        PeakWorkingSetDeltaMB = [math]::Round([math]::Max(0L, [long]($afterPeakWorkingSet - $beforePeakWorkingSet)) / 1MB, 3)
        ManagedDeltaMB    = [math]::Round(([GC]::GetTotalMemory($false) - $beforeManaged) / 1MB, 3)
        WorkbookValidationStatus = $workbookValidation.Status
        WorkbookOpenStatus = $workbookValidation.OpenStatus
        WorkbookOpenXmlStatus = $workbookValidation.OpenXmlStatus
        WorkbookValidationMs = $workbookValidation.Milliseconds
        WorkbookValidationError = $workbookValidation.Error
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

function Get-BenchmarkModuleDetails {
    param([string] $Name)

    $module = Get-Module $Name | Select-Object -First 1
    if (-not $module) {
        return $null
    }

    $prerelease = $null
    $psData = $null
    if ($module.PrivateData -and $module.PrivateData.PSData) {
        $psData = $module.PrivateData.PSData
    }

    if ($psData -is [Collections.IDictionary]) {
        if ($psData.Contains('Prerelease') -and $psData['Prerelease']) {
            $prerelease = [string]$psData['Prerelease']
        }
    } elseif ($psData -and $psData.PSObject.Properties['Prerelease'] -and $psData.Prerelease) {
        $prerelease = [string]$psData.Prerelease
    }

    $version = $module.Version.ToString()
    [pscustomobject]@{
        Name = $module.Name
        Version = $version
        Prerelease = $prerelease
        DisplayVersion = if ($prerelease) { "$version-$prerelease" } else { $version }
        ModuleBase = $module.ModuleBase
        Path = $module.Path
    }
}

function Get-LoadedAssemblyDetails {
    param([string] $Name)

    $assembly = [AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GetName().Name -eq $Name } |
        Sort-Object { $_.GetName().Version } -Descending |
        Select-Object -First 1

    if (-not $assembly) {
        return $null
    }

    [pscustomobject]@{
        Name = $assembly.GetName().Name
        Version = $assembly.GetName().Version.ToString()
        Location = $assembly.Location
        FullName = $assembly.FullName
    }
}

function Get-BenchmarkEnvironment {
    $processor = $null
    $computerSystem = $null
    if (Get-Command Get-CimInstance -ErrorAction SilentlyContinue) {
        try {
            $processor = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
            $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop | Select-Object -First 1
        } catch {
            $processor = $null
            $computerSystem = $null
        }
    }

    [pscustomobject]@{
        MachineName = [Environment]::MachineName
        OSDescription = [Runtime.InteropServices.RuntimeInformation]::OSDescription
        OSArchitecture = [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
        ProcessArchitecture = [Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString()
        DotNetVersion = [Environment]::Version.ToString()
        ProcessorName = if ($processor) { $processor.Name.Trim() } else { $null }
        ProcessorCores = if ($processor) { [int]$processor.NumberOfCores } else { $null }
        ProcessorLogicalProcessors = if ($processor) { [int]$processor.NumberOfLogicalProcessors } else { $null }
        TotalPhysicalMemoryGB = if ($computerSystem) { [math]::Round([double]$computerSystem.TotalPhysicalMemory / 1GB, 2) } else { $null }
    }
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
                        MedianWorkingSetBeforeMB = [double]$_.MedianWorkingSetBeforeMB
                        MedianWorkingSetAfterMB = [double]$_.MedianWorkingSetAfterMB
                        MedianWorkingSetDeltaMB = [double]$_.MedianWorkingSetDeltaMB
                        MedianPeakWorkingSetMB = [double]$_.MedianPeakWorkingSetMB
                        MedianPeakWorkingSetDeltaMB = [double]$_.MedianPeakWorkingSetDeltaMB
                        MedianManagedDeltaMB = [double]$_.MedianManagedDeltaMB
                        WorkbookValidationStatus = if ([int]$_.WorkbookValidationFailed -gt 0) {
                            'failed'
                        } elseif ([int]$_.WorkbookValidationPassed -gt 0) {
                            'passed'
                        } else {
                            'skipped'
                        }
                        WorkbookValidationPassed = [int]$_.WorkbookValidationPassed
                        WorkbookValidationFailed = [int]$_.WorkbookValidationFailed
                        WorkbookValidationSkipped = [int]$_.WorkbookValidationSkipped
                        MedianWorkbookValidationMs = [double]$_.MedianWorkbookValidationMs
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

        foreach ($engineName in @('PSWriteOffice', 'ImportExcel', 'ExcelFast', 'NativeCsv', 'CsvHelper')) {
            $engineResult = @($engineResults | Where-Object Engine -eq $engineName | Select-Object -First 1)
            $engineSummary = @($group.Group | Where-Object Engine -eq $engineName | Select-Object -First 1)
            $prefix = $engineName -replace '[^A-Za-z0-9]', ''
            $row["${prefix}Status"] = if ($engineResult.Count -gt 0) {
                'tested'
            } elseif ($engineSummary.Count -gt 0) {
                'failed'
            } elseif ($Engine -contains $engineName) {
                'not supported by scenario'
            } else {
                'not selected'
            }
            $row["${prefix}Ms"] = if ($engineResult.Count -gt 0) { $engineResult[0].MedianMs } else { $null }
            $row["${prefix}VsFastest"] = if ($engineResult.Count -gt 0) { $engineResult[0].TimeVsFastest } else { $null }
            $row["${prefix}FileKB"] = if ($engineResult.Count -gt 0) { $engineResult[0].MedianFileKB } else { $null }
            $row["${prefix}WorkingSetDeltaMB"] = if ($engineResult.Count -gt 0) { $engineResult[0].MedianWorkingSetDeltaMB } else { $null }
            $row["${prefix}PeakWorkingSetDeltaMB"] = if ($engineResult.Count -gt 0) { $engineResult[0].MedianPeakWorkingSetDeltaMB } else { $null }
            $row["${prefix}ManagedDeltaMB"] = if ($engineResult.Count -gt 0) { $engineResult[0].MedianManagedDeltaMB } else { $null }
            $row["${prefix}WorkbookValidationStatus"] = if ($engineResult.Count -gt 0) { $engineResult[0].WorkbookValidationStatus } else { $null }
            $row["${prefix}WorkbookValidationMs"] = if ($engineResult.Count -gt 0) { $engineResult[0].MedianWorkbookValidationMs } else { $null }
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
$repeatCountExplicit = $PSBoundParameters.ContainsKey('RepeatCount') -and [int]$PSBoundParameters['RepeatCount'] -gt 0

function Test-CsvMicroBenchmarkScenario {
    param([object] $ScenarioObject)

    $key = [string]$ScenarioObject.Key
    return $key -eq 'csv-read' -or
        $key -eq 'csv-read-source' -or
        $key -eq 'csv-write'
}

function Get-CsvMicroBenchmarkRepeatCount {
    param(
        [string] $SuiteName,
        [int] $BaseRepeatCount
    )

    $minimum = switch ($SuiteName) {
        'Smoke' { 11 }
        'Standard' { 51 }
        'Full' { 51 }
        'Large' { 11 }
        'SuperLarge' { 3 }
    }

    [math]::Max($BaseRepeatCount, $minimum)
}

function Get-BenchmarkScenarioIdentity {
    param([object] $ScenarioObject)

    '{0}|{1}|{2}|{3}' -f $ScenarioObject.Engine, $ScenarioObject.Key, $ScenarioObject.Profile, $ScenarioObject.FileStem
}

function Get-BenchmarkScenarioRepeatCount {
    param([object] $ScenarioObject)

    if ($repeatCountExplicit) {
        return $RepeatCount
    }

    if (Test-CsvMicroBenchmarkScenario -ScenarioObject $ScenarioObject) {
        return Get-CsvMicroBenchmarkRepeatCount -SuiteName $Suite -BaseRepeatCount $RepeatCount
    }

    return $RepeatCount
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
                FileExtension = $_.FileExtension
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

            if ($SkipFollowUps.IsPresent) {
                return $false
            }

            $matchingFollowUps = @(Get-SelectedFollowUps -ScenarioObject $_)
            return $matchingFollowUps.Count -gt 0
        }
)

if ($selectedScenarios.Count -eq 0) {
    throw 'No benchmark scenarios matched the requested suite, engine, and scenario filters.'
}

$null = New-Item -ItemType Directory -Force -Path $moduleRoot, $workRoot

$selectedEngines = @($selectedScenarios | Select-Object -ExpandProperty Engine -Unique)

if ($selectedEngines -contains 'ImportExcel') {
    Ensure-ImportExcel
}
if ($selectedEngines -contains 'ExcelFast') {
    Ensure-ExcelFast
}
$requiresCsvHelper = @($selectedScenarios | Where-Object { $_.Engine -eq 'CsvHelper' }).Count -gt 0
if ($requiresCsvHelper) {
    Ensure-CsvHelper
}

if (-not ($Engine -contains 'ExcelFast')) {
    $selectedScenarios = @($selectedScenarios | Where-Object { $_.Engine -ne 'ExcelFast' })
    if ($selectedScenarios.Count -eq 0) {
        throw 'No benchmark scenarios matched after removing unavailable ExcelFast.'
    }
}

$selectedEngines = @($selectedScenarios | Select-Object -ExpandProperty Engine -Unique)

$requiresPSWriteOfficeModule = @($selectedScenarios | Where-Object { $_.Engine -eq 'PSWriteOffice' }).Count -gt 0
$requiresWorkbookValidation = (-not $SkipWorkbookValidation.IsPresent) -and
    @($selectedScenarios | Where-Object { [bool]$_.ValidateWorkbook }).Count -gt 0

if ($requiresPSWriteOfficeModule -or $requiresWorkbookValidation) {
    if (-not [string]::IsNullOrWhiteSpace($OfficeIMORoot)) {
        $env:OfficeIMORoot = $OfficeIMORoot
    } elseif (-not $env:OfficeIMORoot) {
        $env:OfficeIMORoot = Join-Path $repoRoot '.missing-officeimo'
    }

    if (-not $SkipPSWriteOfficeBuild.IsPresent) {
        Invoke-PSWriteOfficeBuild -Configuration $PSWriteOfficeConfiguration
    }

    $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
    $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = $PSWriteOfficeConfiguration
    Import-Module (Join-Path $repoRoot 'PSWriteOffice.psd1') -Force -ErrorAction Stop
}
if ($selectedEngines -contains 'ImportExcel') {
    Import-Module ImportExcel -Force -ErrorAction Stop
}
if ($selectedEngines -contains 'ExcelFast') {
    Import-Module ExcelFast -Force -ErrorAction Stop
}

$scenarioRepeatCounts = @{}
foreach ($benchmarkScenario in $selectedScenarios) {
    $scenarioRepeatCounts[(Get-BenchmarkScenarioIdentity -ScenarioObject $benchmarkScenario)] = Get-BenchmarkScenarioRepeatCount -ScenarioObject $benchmarkScenario
}

$maxRepeatCount = if ($scenarioRepeatCounts.Count -gt 0) {
    [int](($scenarioRepeatCounts.Values | Measure-Object -Maximum).Maximum)
} else {
    $RepeatCount
}

$results = [Collections.Generic.List[object]]::new()
foreach ($rows in $RowCount) {
    $profileCache = @{}
    for ($iteration = 1; $iteration -le $maxRepeatCount; $iteration++) {
        foreach ($benchmarkScenario in $selectedScenarios) {
            $scenarioRepeatCount = [int]$scenarioRepeatCounts[(Get-BenchmarkScenarioIdentity -ScenarioObject $benchmarkScenario)]
            if ($iteration -gt $scenarioRepeatCount) {
                continue
            }

            if (-not $profileCache.ContainsKey($benchmarkScenario.Profile)) {
                $profileCache[$benchmarkScenario.Profile] = Get-BenchmarkData -Profile $benchmarkScenario.Profile -Count $rows
            }

            $profile = $profileCache[$benchmarkScenario.Profile]
            $extension = if ($benchmarkScenario.FileExtension) { $benchmarkScenario.FileExtension } else { 'xlsx' }
            $path = Join-Path $workRoot ('{0}-{1}-{2}.{3}' -f $benchmarkScenario.FileStem, $rows, $iteration, $extension.TrimStart('.'))
            $sourcePath = Join-Path $workRoot ('{0}-{1}-{2}.source.csv' -f $benchmarkScenario.FileStem, $rows, $iteration)
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
                SourcePath = $sourcePath
                Range = 'A1:{0}{1}' -f $rangeEndColumn, ($rows + 1)
                RangeEndCell = '{0}{1}' -f $rangeEndColumn, ($rows + 1)
            }

            if (Test-Path $context.Path) {
                Remove-Item $context.Path -Force
            }
            if (Test-Path $context.SourcePath) {
                Remove-Item $context.SourcePath -Force
            }

            if ($benchmarkScenario.Setup) {
                & $benchmarkScenario.Setup $context
            }

            $validateWorkbook = (-not $SkipWorkbookValidation.IsPresent) -and [bool]$benchmarkScenario.ValidateWorkbook
            $results.Add((Invoke-BenchmarkOperation -Context $context -ScenarioKey $benchmarkScenario.Key -ScenarioName $benchmarkScenario.Name -ScriptBlock $benchmarkScenario.Script -ValidateWorkbook $validateWorkbook))

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
            MedianWorkingSetBeforeMB = if ($passed.Count) { [math]::Round((Get-MedianValue -InputObject $passed -PropertyName WorkingSetBeforeMB), 3) } else { 0 }
            MedianWorkingSetAfterMB = if ($passed.Count) { [math]::Round((Get-MedianValue -InputObject $passed -PropertyName WorkingSetAfterMB), 3) } else { 0 }
            MedianWorkingSetDeltaMB = if ($passed.Count) { [math]::Round((Get-MedianValue -InputObject $passed -PropertyName WorkingSetDeltaMB), 3) } else { 0 }
            MedianPeakWorkingSetMB = if ($passed.Count) { [math]::Round((Get-MedianValue -InputObject $passed -PropertyName PeakWorkingSetMB), 3) } else { 0 }
            MedianPeakWorkingSetDeltaMB = if ($passed.Count) { [math]::Round((Get-MedianValue -InputObject $passed -PropertyName PeakWorkingSetDeltaMB), 3) } else { 0 }
            MedianManagedDeltaMB = if ($passed.Count) { [math]::Round((Get-MedianValue -InputObject $passed -PropertyName ManagedDeltaMB), 3) } else { 0 }
            WorkbookValidationPassed = @($passed | Where-Object WorkbookValidationStatus -eq 'Passed').Count
            WorkbookValidationFailed = @($passed | Where-Object WorkbookValidationStatus -eq 'Failed').Count
            WorkbookValidationSkipped = @($passed | Where-Object WorkbookValidationStatus -eq 'Skipped').Count
            MedianWorkbookValidationMs = if ($passed.Count) { [math]::Round((Get-MedianValue -InputObject $passed -PropertyName WorkbookValidationMs), 3) } else { 0 }
        }
    } |
    Sort-Object ScenarioKey, Profile, Rows, Engine

$summary | Export-Csv -NoTypeInformation -Path $summaryPath
$comparison = @(New-BenchmarkComparison -Summary $summary)
$comparison |
    Select-Object ScenarioKey, Scenario, Profile, Rows, FastestEngine, FastestMs, PSWriteOfficeStatus, PSWriteOfficeMs, PSWriteOfficeRank, PSWriteOfficeVsFastest, PSWriteOfficeVsFastestText, LeadText, Rating, ImportExcelStatus, ImportExcelMs, ImportExcelVsFastest, ExcelFastStatus, ExcelFastMs, ExcelFastVsFastest, NativeCsvStatus, NativeCsvMs, NativeCsvVsFastest, CsvHelperStatus, CsvHelperMs, CsvHelperVsFastest, SmallestFileEngine, PSWriteOfficeFileKB, ImportExcelFileKB, ExcelFastFileKB, NativeCsvFileKB, CsvHelperFileKB, PSWriteOfficeWorkbookValidationStatus, ImportExcelWorkbookValidationStatus, ExcelFastWorkbookValidationStatus, NativeCsvWorkbookValidationStatus, CsvHelperWorkbookValidationStatus, PSWriteOfficeWorkbookValidationMs, ImportExcelWorkbookValidationMs, ExcelFastWorkbookValidationMs, NativeCsvWorkbookValidationMs, CsvHelperWorkbookValidationMs, PSWriteOfficeWorkingSetDeltaMB, ImportExcelWorkingSetDeltaMB, ExcelFastWorkingSetDeltaMB, NativeCsvWorkingSetDeltaMB, CsvHelperWorkingSetDeltaMB, PSWriteOfficePeakWorkingSetDeltaMB, ImportExcelPeakWorkingSetDeltaMB, ExcelFastPeakWorkingSetDeltaMB, NativeCsvPeakWorkingSetDeltaMB, CsvHelperPeakWorkingSetDeltaMB, PSWriteOfficeManagedDeltaMB, ImportExcelManagedDeltaMB, ExcelFastManagedDeltaMB, NativeCsvManagedDeltaMB, CsvHelperManagedDeltaMB |
    Export-Csv -NoTypeInformation -Path $comparisonCsvPath
$comparison | ConvertTo-Json -Depth 8 | Set-Content -Path $comparisonJsonPath -Encoding UTF8
$officeIMOExcelAssemblyPath = Join-Path (Join-Path (Join-Path (Join-Path (Join-Path (Join-Path $repoRoot 'Sources') 'PSWriteOffice') 'bin') $PSWriteOfficeConfiguration) 'net8.0') 'OfficeIMO.Excel.dll'
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

$moduleDetails = [ordered]@{
    PSWriteOffice = Get-BenchmarkModuleDetails -Name PSWriteOffice
    ImportExcel = Get-BenchmarkModuleDetails -Name ImportExcel
    ExcelFast = Get-BenchmarkModuleDetails -Name ExcelFast
    CsvHelper = Get-LoadedAssemblyDetails -Name CsvHelper
    NativeCsv = [pscustomobject]@{
        Name = 'Microsoft.PowerShell.Utility'
        Version = $PSVersionTable.PSVersion.ToString()
        Prerelease = $null
        DisplayVersion = $PSVersionTable.PSVersion.ToString()
        ModuleBase = $null
        Path = $null
    }
}
$assemblyDetails = [ordered]@{
    OfficeOpenXml = Get-LoadedAssemblyDetails -Name OfficeOpenXml
    OfficeIMOExcel = Get-LoadedAssemblyDetails -Name OfficeIMO.Excel
    CsvHelper = Get-LoadedAssemblyDetails -Name CsvHelper
}

[pscustomobject]@{
    PowerShellVersion = $PSVersionTable.PSVersion.ToString()
    PSEdition = $PSEdition
    Environment = Get-BenchmarkEnvironment
    Suite = $Suite
    ImportExcel = if ($moduleDetails.ImportExcel) { $moduleDetails.ImportExcel.DisplayVersion } else { $null }
    ExcelFast = if ($moduleDetails.ExcelFast) { $moduleDetails.ExcelFast.DisplayVersion } else { $null }
    CsvHelper = if ($moduleDetails.CsvHelper) { $moduleDetails.CsvHelper.Version } else { $null }
    NativeCsv = $moduleDetails.NativeCsv.DisplayVersion
    PSWriteOffice = if ($moduleDetails.PSWriteOffice) { $moduleDetails.PSWriteOffice.DisplayVersion } else { $null }
    Modules = $moduleDetails
    Assemblies = $assemblyDetails
    OfficeIMOExcelAssembly = $officeIMOExcelAssemblyVersion
    OfficeIMOExcelAssemblyPath = if (Test-Path $officeIMOExcelAssemblyPath) { $officeIMOExcelAssemblyPath } else { $null }
    Engines = $Engine
    ScenarioFilter = $Scenario
    RowCount = $RowCount
    RepeatCount = $RepeatCount
    RepeatCountExplicit = $repeatCountExplicit
    MaxRepeatCount = $maxRepeatCount
    ScenarioRepeatPolicy = @(
        $selectedScenarios |
            Sort-Object Key, Engine, Profile |
            ForEach-Object {
                [pscustomobject]@{
                    Key = $_.Key
                    Engine = $_.Engine
                    Profile = $_.Profile
                    RepeatCount = [int]$scenarioRepeatCounts[(Get-BenchmarkScenarioIdentity -ScenarioObject $_)]
                    Reason = if (-not $repeatCountExplicit -and (Test-CsvMicroBenchmarkScenario -ScenarioObject $_)) {
                        'csv microbenchmark default'
                    } elseif ($repeatCountExplicit) {
                        'explicit RepeatCount'
                    } else {
                        'suite default'
                    }
                }
            }
    )
    SkipFollowUps = $SkipFollowUps.IsPresent
    SkipWorkbookValidation = $SkipWorkbookValidation.IsPresent
    ScenarioCount = $selectedScenarios.Count
    RepoRoot = $repoRoot.Path
    OutputDirectory = $OutputDirectory
    WorkRoot = $workRoot
    ModuleCache = $moduleRoot
    OfficeIMORoot = $env:OfficeIMORoot
    PSWriteOfficeUseDevelopmentBinaries = $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES
    PSWriteOfficeDevelopmentConfiguration = $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION
    PSWriteOfficeBuildSkipped = $SkipPSWriteOfficeBuild.IsPresent
    ResultsPath = $resultsPath
    SummaryPath = $summaryPath
    ComparisonCsvPath = $comparisonCsvPath
    ComparisonJsonPath = $comparisonJsonPath
} | ConvertTo-Json -Depth 8 | Set-Content -Path $metadataPath -Encoding UTF8

Write-Host "Results: $resultsPath"
Write-Host "Summary: $summaryPath"
Write-Host "Comparison CSV: $comparisonCsvPath"
Write-Host "Comparison JSON: $comparisonJsonPath"
Write-Host "Metadata: $metadataPath"
$comparison |
    Select-Object ScenarioKey, Profile, Rows, FastestEngine, FastestMs, PSWriteOfficeMs, PSWriteOfficeVsFastestText, Rating |
    Format-Table -AutoSize
$summary | Format-Table -AutoSize
