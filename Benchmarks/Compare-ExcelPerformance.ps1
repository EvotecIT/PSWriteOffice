param(
    [int[]] $RowCount = @(1000, 10000, 25000),
    [int] $RepeatCount = 3,
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\Ignore\Benchmarks\ExcelPerformance'),
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
$workRoot = Join-Path $OutputDirectory ('Run-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Force -Path $moduleRoot, $workRoot

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

function Get-RowCount {
    param([object] $Rows)

    if ($null -eq $Rows) { return 0 }
    if ($Rows -is [array]) { return $Rows.Count }
    return @($Rows).Count
}

function Invoke-BenchmarkOperation {
    param(
        [string] $Engine,
        [string] $Scenario,
        [int] $Rows,
        [int] $Iteration,
        [scriptblock] $ScriptBlock,
        [string] $FilePath
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
        $result = & $ScriptBlock
        $resultCount = Get-RowCount -Rows $result
    } catch {
        $status = 'Failed'
        $errorMessage = $_.Exception.Message
    }
    $stopwatch.Stop()
    $process.Refresh()

    [pscustomobject]@{
        TimestampUtc      = [datetime]::UtcNow.ToString('o')
        Engine            = $Engine
        Scenario          = $Scenario
        Rows              = $Rows
        Iteration         = $Iteration
        Milliseconds      = [math]::Round($stopwatch.Elapsed.TotalMilliseconds, 3)
        ResultCount       = $resultCount
        FileBytes         = if (Test-Path $FilePath) { (Get-Item $FilePath).Length } else { 0 }
        WorkingSetDeltaMB = [math]::Round(($process.WorkingSet64 - $beforeWorkingSet) / 1MB, 3)
        ManagedDeltaMB    = [math]::Round(([GC]::GetTotalMemory($false) - $beforeManaged) / 1MB, 3)
        Status            = $status
        Error             = $errorMessage
    }
}

Ensure-ImportExcel
Ensure-ExcelFast

$env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
$env:OfficeIMORoot = Join-Path $repoRoot '.missing-officeimo'
Import-Module (Join-Path $repoRoot 'PSWriteOffice.psd1') -Force -ErrorAction Stop
Import-Module ImportExcel -Force -ErrorAction Stop
Import-Module ExcelFast -Force -ErrorAction Stop

$results = [Collections.Generic.List[object]]::new()
foreach ($rows in $RowCount) {
    $data = @(New-BenchmarkRows -Count $rows)
    for ($iteration = 1; $iteration -le $RepeatCount; $iteration++) {
        $scenarios = @(
            @{
                Engine = 'PSWriteOffice'
                Name = 'Export objects as table'
                Path = Join-Path $workRoot "pswriteoffice-objects-table-$rows-$iteration.xlsx"
                Script = { $data | Export-OfficeExcel -Path $scenarios[$scenarioIndex].Path -WorksheetName Data -TableName Data }
            },
            @{
                Engine = 'ImportExcel'
                Name = 'Export objects as table'
                Path = Join-Path $workRoot "importexcel-objects-table-$rows-$iteration.xlsx"
                Script = { $data | Export-Excel -Path $scenarios[$scenarioIndex].Path -WorksheetName Data -TableName Data -AutoFilter }
            },
            @{
                Engine = 'PSWriteOffice'
                Name = 'Export objects default'
                Path = Join-Path $workRoot "pswriteoffice-objects-default-$rows-$iteration.xlsx"
                Script = { $data | Export-OfficeExcel -Path $scenarios[$scenarioIndex].Path -WorksheetName Data }
            },
            @{
                Engine = 'ImportExcel'
                Name = 'Export objects default'
                Path = Join-Path $workRoot "importexcel-objects-default-$rows-$iteration.xlsx"
                Script = { $data | Export-Excel -Path $scenarios[$scenarioIndex].Path -WorksheetName Data }
            },
            @{
                Engine = 'ExcelFast'
                Name = 'Export objects default'
                Path = Join-Path $workRoot "excelfast-objects-default-$rows-$iteration.xlsx"
                Script = { Export-Workbook -Destination $scenarios[$scenarioIndex].Path -InputObject $data -SheetName Data -Force }
            },
            @{
                Engine = 'PSWriteOffice'
                Name = 'Export objects no table'
                Path = Join-Path $workRoot "pswriteoffice-objects-notable-$rows-$iteration.xlsx"
                Script = { $data | Export-OfficeExcel -Path $scenarios[$scenarioIndex].Path -WorksheetName Data -NoTable }
            },
            @{
                Engine = 'ImportExcel'
                Name = 'Export objects no table'
                Path = Join-Path $workRoot "importexcel-objects-notable-$rows-$iteration.xlsx"
                Script = { $data | Export-Excel -Path $scenarios[$scenarioIndex].Path -WorksheetName Data }
            },
            @{
                Engine = 'PSWriteOffice'
                Name = 'Export objects table autofit'
                Path = Join-Path $workRoot "pswriteoffice-objects-table-autofit-$rows-$iteration.xlsx"
                Script = { $data | Export-OfficeExcel -Path $scenarios[$scenarioIndex].Path -WorksheetName Data -TableName Data -AutoFit }
            },
            @{
                Engine = 'ImportExcel'
                Name = 'Export objects table autofit'
                Path = Join-Path $workRoot "importexcel-objects-table-autofit-$rows-$iteration.xlsx"
                Script = { $data | Export-Excel -Path $scenarios[$scenarioIndex].Path -WorksheetName Data -TableName Data -AutoFilter -AutoSize }
            }
        )

        for ($scenarioIndex = 0; $scenarioIndex -lt $scenarios.Count; $scenarioIndex++) {
            $scenario = $scenarios[$scenarioIndex]
            if (Test-Path $scenario.Path) {
                Remove-Item $scenario.Path -Force
            }
            $results.Add((Invoke-BenchmarkOperation -Engine $scenario.Engine -Scenario $scenario.Name -Rows $rows -Iteration $iteration -FilePath $scenario.Path -ScriptBlock $scenario.Script))

            if ($scenario.Name -eq 'Export objects as table' -and (Test-Path $scenario.Path)) {
                $importName = 'Import full sheet from table export'
                $importPath = $scenario.Path
                if ($scenario.Engine -eq 'PSWriteOffice') {
                    $results.Add((Invoke-BenchmarkOperation -Engine 'PSWriteOffice' -Scenario $importName -Rows $rows -Iteration $iteration -FilePath $importPath -ScriptBlock { Import-OfficeExcel -Path $importPath -WorksheetName Data }))
                } elseif ($scenario.Engine -eq 'ImportExcel') {
                    $results.Add((Invoke-BenchmarkOperation -Engine 'ImportExcel' -Scenario $importName -Rows $rows -Iteration $iteration -FilePath $importPath -ScriptBlock { Import-Excel -Path $importPath -WorksheetName Data }))
                }
            } elseif ($scenario.Name -eq 'Export objects default' -and (Test-Path $scenario.Path)) {
                $importName = 'Import full sheet from default export'
                $importPath = $scenario.Path
                if ($scenario.Engine -eq 'PSWriteOffice') {
                    $results.Add((Invoke-BenchmarkOperation -Engine 'PSWriteOffice' -Scenario $importName -Rows $rows -Iteration $iteration -FilePath $importPath -ScriptBlock { Import-OfficeExcel -Path $importPath -WorksheetName Data }))
                } elseif ($scenario.Engine -eq 'ImportExcel') {
                    $results.Add((Invoke-BenchmarkOperation -Engine 'ImportExcel' -Scenario $importName -Rows $rows -Iteration $iteration -FilePath $importPath -ScriptBlock { Import-Excel -Path $importPath -WorksheetName Data }))
                } elseif ($scenario.Engine -eq 'ExcelFast') {
                    $results.Add((Invoke-BenchmarkOperation -Engine 'ExcelFast' -Scenario $importName -Rows $rows -Iteration $iteration -FilePath $importPath -ScriptBlock { Import-Workbook -Path $importPath -SheetName Data }))
                }
            }
        }
    }
}

$resultsPath = Join-Path $workRoot 'excel-performance-results.csv'
$summaryPath = Join-Path $workRoot 'excel-performance-summary.csv'
$metadataPath = Join-Path $workRoot 'metadata.json'

$results | Export-Csv -NoTypeInformation -Path $resultsPath
$summary = $results |
    Group-Object Engine, Scenario, Rows |
    ForEach-Object {
        $passed = @($_.Group | Where-Object Status -eq 'Passed')
        $ordered = @($passed | Sort-Object Milliseconds)
        $median = if ($ordered.Count -eq 0) { 0 } else { $ordered[[int][math]::Floor(($ordered.Count - 1) / 2)].Milliseconds }
        [pscustomobject]@{
            Engine       = $_.Group[0].Engine
            Scenario     = $_.Group[0].Scenario
            Rows         = $_.Group[0].Rows
            Runs         = $_.Group.Count
            Passed       = $passed.Count
            MedianMs     = $median
            MinMs        = if ($passed.Count) { ($passed | Measure-Object Milliseconds -Minimum).Minimum } else { 0 }
            MaxMs        = if ($passed.Count) { ($passed | Measure-Object Milliseconds -Maximum).Maximum } else { 0 }
            MedianFileKB = if ($passed.Count) { [math]::Round((($passed | Sort-Object FileBytes)[[int][math]::Floor(($passed.Count - 1) / 2)].FileBytes) / 1KB, 1) } else { 0 }
        }
    } |
    Sort-Object Scenario, Rows, Engine

$summary | Export-Csv -NoTypeInformation -Path $summaryPath
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
    ImportExcel = (Get-Module ImportExcel).Version.ToString()
    ExcelFast = (Get-Module ExcelFast).Version.ToString()
    PSWriteOffice = (Get-Module PSWriteOffice).Version.ToString()
    OfficeIMOExcelAssembly = $officeIMOExcelAssemblyVersion
    RowCount = $RowCount
    RepeatCount = $RepeatCount
    ResultsPath = $resultsPath
    SummaryPath = $summaryPath
} | ConvertTo-Json -Depth 5 | Set-Content -Path $metadataPath -Encoding UTF8

Write-Host "Results: $resultsPath"
Write-Host "Summary: $summaryPath"
Write-Host "Metadata: $metadataPath"
$summary | Format-Table -AutoSize
