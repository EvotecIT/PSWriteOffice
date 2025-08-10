Describe 'Export-OfficeExcel cmdlet' {
    It 'creates an Excel file with a worksheet and table' {
        $path = Join-Path $TestDrive 'test.xlsx'
        $data = 1..3 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
        $data | Export-OfficeExcel -FilePath $path -WorksheetName 'Data'
        Test-Path $path | Should -BeTrue
    }

    It 'appends to an existing table when Append is used' {
        $path = Join-Path $TestDrive 'append.xlsx'
        $first = 1..2 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
        $first | Export-OfficeExcel -FilePath $path -WorksheetName 'Data'
        $second = 3..4 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
        $second | Export-OfficeExcel -FilePath $path -WorksheetName 'Data' -Append
        $rows = Import-OfficeExcel -FilePath $path -WorkSheetName 'Data'
        $rows.Count | Should -Be 4
        $rows[-1].Value | Should -Be 4
    }

    It 'includes all properties when AllProperties is used' {
        $path = Join-Path $TestDrive 'allprops.xlsx'
        $data = @(
            [PSCustomObject]@{ First = 1; Second = 'A' },
            [PSCustomObject]@{ First = 2 }
        )
        $data | Export-OfficeExcel -FilePath $path -AllProperties
        $rows = Import-OfficeExcel -FilePath $path
        # Check that both rows have the Second property
        ($rows[0].PSObject.Properties.Name -contains 'Second') | Should -BeTrue
        ($rows[1].PSObject.Properties.Name -contains 'Second') | Should -BeTrue
        $rows[0].Second | Should -Be 'A'
        $rows[1].Second | Should -BeNullOrEmpty
    }

    It 'auto sizes columns and freezes panes when switches are used' {
        $path = Join-Path $TestDrive 'format.xlsx'
        $data = 1..2 | ForEach-Object { [PSCustomObject]@{ Name = "Row$_"; Value = "Some very long value $_" } }
        $data | Export-OfficeExcel -FilePath $path -WorksheetName 'Data' -AutoSize -FreezeTopRow -FreezeFirstColumn
        $dll = Join-Path $PSScriptRoot '..' 'Sources' 'PSWriteOffice' 'bin' 'Debug' 'net8.0' 'ClosedXML.dll'
        Add-Type -Path $dll
        $wb = [ClosedXML.Excel.XLWorkbook]::new($path)
        $ws = $wb.Worksheet('Data')
        # AutoSize adjusts columns to content - Name column is narrow, Value column is wide
        $ws.Column(1).Width | Should -Not -Be 8.43  # Should be different from default
        $ws.Column(2).Width | Should -BeGreaterThan 8.43  # Long content should be wider
        $ws.SheetView.SplitRow | Should -Be 1
        $ws.SheetView.SplitColumn | Should -Be 1
    }

    It 'creates file even when path does not exist' {
        $data = 1..3 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
        $path = Join-Path $TestDrive 'newfile.xlsx'
        $data | Export-OfficeExcel -FilePath $path
        Test-Path $path | Should -BeTrue
    }

    It 'creates a pivot table when definition is provided' {
        $path = Join-Path $TestDrive 'pivot.xlsx'
        $data = @(
            [PSCustomObject]@{ Category = 'A'; Value = 1 }
            [PSCustomObject]@{ Category = 'A'; Value = 2 }
            [PSCustomObject]@{ Category = 'B'; Value = 3 }
        )

        $pivot = @{ Name = 'Pivot1'; SourceRange = 'A1:B4'; TargetCell = 'D2'; RowFields = @('Category'); Values = @{ Value = 'Sum' } }
        $data | Export-OfficeExcel -FilePath $path -WorksheetName 'Data' -PivotTables $pivot

        $dll = Join-Path $PSScriptRoot '..' 'Sources' 'PSWriteOffice' 'bin' 'Debug' 'net8.0' 'ClosedXML.dll'
        Add-Type -Path $dll
        $wb = [ClosedXML.Excel.XLWorkbook]::new($path)
        $ws = $wb.Worksheet('Data')
        $ws.PivotTables.Count | Should -BeGreaterThan 0
    }

    It 'adds a chart when chart specification is provided' {
        $path = Join-Path $TestDrive 'chart.xlsx'
        $data = 1..3 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
        $chart = @{ Title = 'Chart1'; Range = 'A1:B3' }
        $data | Export-OfficeExcel -FilePath $path -WorksheetName 'Data' -Charts $chart

        $openXml = Join-Path $PSScriptRoot '..' 'Sources' 'PSWriteOffice' 'bin' 'Debug' 'net8.0' 'DocumentFormat.OpenXml.dll'
        Add-Type -Path $openXml
        $doc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($path, $false)
        $sheet = $doc.WorkbookPart.Workbook.Sheets.ChildElements | Where-Object { $_.Name -eq 'Data' }
        $wsPart = $doc.WorkbookPart.GetPartById($sheet.Id)
        $wsPart.DrawingsPart.ChartParts.Count | Should -BeGreaterThan 0
        $doc.Close()
    }
}
