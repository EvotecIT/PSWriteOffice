function Get-ExcelBenchmarkRegionGroups {
    param([object[]] $Rows)

    foreach ($region in @('NA', 'EU', 'APAC', 'LATAM')) {
        [pscustomobject]@{
            Name = $region
            TableName = 'Data_' + $region
            Data = @($Rows | Where-Object { $_.Region -eq $region })
        }
    }
}

function Get-ExcelBenchmarkSummaryRows {
    param([int] $Rows)

    @(
        [pscustomobject]@{ Metric = 'Rows'; Formula = 'COUNTA(Data!A2:A{0})' -f ($Rows + 1); NumberFormat = '#,##0' }
        [pscustomobject]@{ Metric = 'Average score'; Formula = 'AVERAGE(Data!G2:G{0})' -f ($Rows + 1); NumberFormat = '#,##0.00' }
        [pscustomobject]@{ Metric = 'Tickets'; Formula = 'SUM(Data!I2:I{0})' -f ($Rows + 1); NumberFormat = '#,##0' }
        [pscustomobject]@{ Metric = 'Enabled'; Formula = 'COUNTIF(Data!E2:E{0},TRUE)' -f ($Rows + 1); NumberFormat = '#,##0' }
    )
}

function Get-ExcelBenchmarkAppendSplit {
    param([object[]] $Rows)

    $initialCount = [math]::Max(1, [math]::Floor($Rows.Count * 0.8))
    [pscustomobject]@{
        Initial = @($Rows | Select-Object -First $initialCount)
        Append = @($Rows | Select-Object -Skip $initialCount)
    }
}

function Get-ExcelBenchmarkSmallSheetGroups {
    param([object[]] $Rows, [int] $SheetCount = 20)

    $safeSheetCount = [math]::Max(1, [math]::Min($SheetCount, [math]::Max(1, $Rows.Count)))
    $rowsPerSheet = [math]::Max(1, [math]::Ceiling($Rows.Count / $safeSheetCount))
    for ($sheetIndex = 0; $sheetIndex -lt $safeSheetCount; $sheetIndex++) {
        $sheetRows = @($Rows | Select-Object -Skip ($sheetIndex * $rowsPerSheet) -First $rowsPerSheet)
        if ($sheetRows.Count -eq 0) { continue }
        [pscustomobject]@{
            Name = 'Sheet{0:00}' -f ($sheetIndex + 1)
            TableName = 'Data{0:00}' -f ($sheetIndex + 1)
            Data = $sheetRows
        }
    }
}

function Get-ExcelBenchmarkWorkbookMergeInput {
    param([object[]] $Rows, [string] $BasePath)

    $firstCount = [math]::Max(1, [math]::Floor($Rows.Count / 2))
    [pscustomobject]@{
        SourceA = [IO.Path]::Combine([IO.Path]::GetDirectoryName($BasePath), ([IO.Path]::GetFileNameWithoutExtension($BasePath) + '.source-a.xlsx'))
        SourceB = [IO.Path]::Combine([IO.Path]::GetDirectoryName($BasePath), ([IO.Path]::GetFileNameWithoutExtension($BasePath) + '.source-b.xlsx'))
        RowsA = @($Rows | Select-Object -First $firstCount)
        RowsB = @($Rows | Select-Object -Skip $firstCount)
    }
}

function Test-ExcelBenchmarkOutput {
    param([object] $Case, [object] $Run)

    if ($Case.OperationKey -in @('WriteCsv', 'CsvToExcel', 'WriteWorkbook')) {
        assertPath $Run.Path
    }

    if ($Case.OperationKey -in @('ReadFullSheet', 'ReadRange', 'ReadNoHeaderRange')) {
        $expectedRows = if ($Case.OperationKey -eq 'ReadNoHeaderRange') {
            [int]$Run.ExpectedRows + 1
        } else {
            [int]$Run.ExpectedRows
        }
        assertValue ([int]$Run.ActualRows) $expectedRows -Message "Expected $expectedRows rows returned by '$($Case.OperationKey)'."
    }

    if ($Case.OperationKey -eq 'ReadUsedRangeDataTable') {
        $expectedRows = [int]$Run.ExpectedRows
        assertValue ([int]$Run.ActualRows) $expectedRows -Message "Expected $expectedRows rows returned by '$($Case.OperationKey)'."
    }

    if ($Case.OperationKey -eq 'ReadTableMetadata') {
        assertValue ([int]$Run.ActualTableCount) 1 -Message 'Expected one workbook table metadata result.'
        assertValue (@($Run.ActualTableNames) -contains 'Data') $true -Message "Expected table metadata to include 'Data'."
    }

    if ($Case.OperationKey -eq 'ReadNamedRangeMetadata') {
        assertValue ([int]$Run.ActualNamedRangeCount) 1 -Message 'Expected one workbook named range metadata result.'
        assertValue (@($Run.ActualNamedRangeNames) -contains 'SalesData') $true -Message "Expected named range metadata to include 'SalesData'."
    }

    if ([bool]$Run.SkipWorkbookValidation) {
        return
    }

    if (-not [bool]$Case.ValidateWorkbook -or [string]$Case.FileExtension -ne '.xlsx') {
        return
    }

    $document = Get-OfficeExcel -Path $Run.Path -ReadOnly
    if ($document) {
        Close-OfficeExcel -Document $document
    }
    Test-ExcelBenchmarkOpenXml -Path $Run.Path
}

function Test-CsvBenchmarkOutput {
    param([object] $Case, [object] $Run)

    $expectedRows = [int]$Run.ExpectedRows
    if ($Case.OperationKey -eq 'ReadCsvSource') {
        assertValue ([int]$Run.ActualRows) $expectedRows -Message "Expected $expectedRows rows returned by '$($Case.OperationKey)'."
        return
    }

    $path = $Run.Path
    assertPath $path
    $actualRows = @(Import-Csv -Path $path).Count
    assertValue $actualRows $expectedRows -Message "Expected $expectedRows rows in '$path'."
}

function Test-ExcelBenchmarkOpenXml {
    param([string] $Path)

    $validatorType = 'DocumentFormat.OpenXml.Validation.OpenXmlValidator' -as [type]
    $spreadsheetType = 'DocumentFormat.OpenXml.Packaging.SpreadsheetDocument' -as [type]
    if ($null -eq $validatorType -or $null -eq $spreadsheetType) {
        return
    }

    $document = $spreadsheetType::Open($Path, $false)
    try {
        $validator = [Activator]::CreateInstance($validatorType)
        $errors = @($validator.Validate($document))
        assertValue $errors.Count 0 -Message "Expected '$Path' to pass OpenXML validation."
    } finally {
        if ($document) {
            $document.Dispose()
        }
    }
}
