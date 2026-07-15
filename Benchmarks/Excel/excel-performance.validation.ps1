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

    if ($Case.OperationKey -in @('WriteCsv', 'WriteCsvGZip', 'CsvToExcel', 'WriteWorkbook')) {
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

    if ([string]$Case.Scenario -in @('objects-default', 'text-objects-default', 'wide-objects-default')) {
        Test-ExcelBenchmarkTabularValues -Case $Case -Run $Run
    }
}

function Test-ExcelBenchmarkTabularValues {
    param([object] $Case, [object] $Run)

    $actualRows = @(Import-OfficeExcel -Path $Run.Path -WorksheetName $Run.WorksheetName)
    $expectedRows = @($Run.Payload)
    assertValue $actualRows.Count $expectedRows.Count -Message "Expected '$($Case.Scenario)' to preserve every data row."
    if ($expectedRows.Count -eq 0) {
        return
    }

    $lastIndex = $expectedRows.Count - 1
    $middleIndex = [int] [Math]::Floor($lastIndex / 2)
    $indexes = @(0, $middleIndex, $lastIndex) | Select-Object -Unique
    foreach ($index in $indexes) {
        $expected = $expectedRows[$index]
        $actual = $actualRows[$index]
        foreach ($property in $expected.PSObject.Properties) {
            $expectedValue = ConvertTo-ExcelBenchmarkComparableValue -Value $property.Value
            $actualProperty = $actual.PSObject.Properties[$property.Name]
            if ($null -eq $actualProperty) {
                throw "Expected '$($Case.Scenario)' row $index to contain column '$($property.Name)'."
            }

            $actualValue = ConvertTo-ExcelBenchmarkComparableValue -Value $actualProperty.Value
            assertValue $actualValue $expectedValue -Message "Expected '$($Case.Scenario)' row $index column '$($property.Name)' to preserve its value."
        }
    }
}

function ConvertTo-ExcelBenchmarkComparableValue {
    param([AllowNull()][object] $Value)

    if ($null -eq $Value -or $Value -is [DBNull]) {
        return '<null>'
    }
    if ($Value -is [datetime]) {
        return $Value.ToString('O', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [bool]) {
        return $Value.ToString().ToLowerInvariant()
    }
    if ($Value -is [double] -or $Value -is [single]) {
        return ([Math]::Round([double] $Value, 10)).ToString('G17', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [IFormattable]) {
        return $Value.ToString($null, [Globalization.CultureInfo]::InvariantCulture)
    }

    return [string] $Value
}

function Test-CsvBenchmarkOutput {
    param([object] $Case, [object] $Run)

    $expectedRows = [int]$Run.ExpectedRows
    if ($Case.OperationKey -in @('ReadCsvSource', 'ReadCsvDataTable', 'ReadCsvGZipDataTable', 'ReadCsvQuickSingleColumn', 'ReadCsvQuickAllColumns')) {
        assertValue ([int]$Run.ActualRows) $expectedRows -Message "Expected $expectedRows rows returned by '$($Case.OperationKey)'."
        if ($Case.OperationKey -eq 'ReadCsvQuickSingleColumn') {
            assertValue ([int]$Run.AccessedFields) $expectedRows -Message "Expected $expectedRows first-column values accessed by '$($Case.OperationKey)'."
            assertValue ([string]$Run.LastValue) ([string]($expectedRows - 1)) -Message "Expected '$($Case.OperationKey)' to access the last Column0 value."
        }
        if ($Case.OperationKey -eq 'ReadCsvQuickAllColumns') {
            $expectedFields = $expectedRows * [int]$Run.ColumnCount
            assertValue ([int]$Run.AccessedFields) $expectedFields -Message "Expected $expectedFields values accessed by '$($Case.OperationKey)'."
            assertValue ([string]$Run.LastValue) ('Value{0}_{1}' -f ($expectedRows - 1), ([int]$Run.ColumnCount - 1)) -Message "Expected '$($Case.OperationKey)' to access the last field value."
        }
        $Run.RowsProcessed = [int]$Run.ActualRows
        return
    }

    $path = $Run.Path
    assertPath $path
    $actualRows = if ($Case.OperationKey -eq 'WriteCsvGZip') {
        $table = ConvertFrom-NativeGZipCsvToDataTable -Path $path
        if ($table -and $table.Rows) { [int]$table.Rows.Count } else { 0 }
    } else {
        @(Import-Csv -Path $path).Count
    }
    assertValue $actualRows $expectedRows -Message "Expected $expectedRows rows in '$path'."
    $Run.RowsProcessed = [int]$actualRows
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
        if ($errors.Count -gt 0) {
            $details = @(
                $errors |
                    Select-Object -First 5 |
                    ForEach-Object {
                        $part = if ($_.Part -and $_.Part.Uri) { [string]$_.Part.Uri } else { 'unknown part' }
                        '{0}: {1}' -f $part, $_.Description
                    }
            ) -join '; '

            throw "Expected '$Path' to pass OpenXML validation. First errors: $details"
        }
    } finally {
        if ($document) {
            $document.Dispose()
        }
    }
}
