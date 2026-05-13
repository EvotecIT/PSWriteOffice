BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Excel DSL surface' {
    It 'creates a workbook with canonical cmdlets' {
        $path = Join-Path $TestDrive 'DslExcel.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Region'
                Set-OfficeExcelCell -Address 'B1' -Value 'Revenue'
                Add-OfficeExcelTable -Data $rows -TableName 'Sales'
            }
        }

        Test-Path $path | Should -BeTrue

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Count | Should -BeGreaterThan 0
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'supports alias-only syntax' {
        $path = Join-Path $TestDrive 'DslExcelAlias.xlsx'
        $rows = @(
            [PSCustomObject]@{ Item = 'Laptop'; Qty = 5 }
            [PSCustomObject]@{ Item = 'Tablet'; Qty = 12 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Orders' -Content {
                ExcelCell -Address 'A1' -Value 'Item'
                ExcelCell -Address 'B1' -Value 'Qty'
                ExcelTable -Data $rows -TableName 'OrdersTable'
            }
        }

        Test-Path $path | Should -BeTrue
    }

    It 'exports and imports objects through operator cmdlets' {
        $path = Join-Path $TestDrive 'ExportOfficeExcel.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100; Internal = 'skip' }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200; Internal = 'skip' }
        )

        $file = $rows |
            Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales' -Title 'Sales Export' -AutoFit -FreezeTopRow -BoldTopRow -ExcludeProperty Internal -PassThru

        $file.FullName | Should -Be $path
        Test-Path $path | Should -BeTrue

        $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales')
        $tables.Count | Should -Be 1
        $tables[0].Range | Should -Be 'A2:B4'

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A2:B4')
        $imported.Count | Should -Be 2
        $imported[0].Region | Should -Be 'NA'
        $imported[0].Revenue | Should -Be 100
        $imported[0].PSObject.Properties.Name | Should -Not -Contain 'Internal'

        { Import-OfficeExcel -Path $path -WorksheetName 'Data' -StartRow 4 -EndRow 2 -StartColumn 1 -EndColumn 2 } |
            Should -Throw '*StartRow must be less than or equal to EndRow*'
        { Import-OfficeExcel -Path $path -WorksheetName 'Data' -StartRow 2 -EndRow 4 -StartColumn 3 -EndColumn 1 } |
            Should -Throw '*StartColumn must be less than or equal to EndColumn*'
    }

    It 'appends rows without rewriting headers' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelAppend.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )
        $moreRows = @(
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 150 }
        )

        $rows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales' -AutoFit
        $moreRows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -Append -TableName 'Sales' -AutoFit

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B4')
        $imported.Count | Should -Be 3
        $imported[2].Region | Should -Be 'APAC'
        $imported[2].Revenue | Should -Be 150

        $hasTableAppend = @([OfficeIMO.Excel.ExcelSheet].GetMethods() | Where-Object Name -eq 'AppendDataTableToTable').Count -gt 0
        if ($hasTableAppend) {
            $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales')
            $tables.Count | Should -Be 1
            $tables[0].Range | Should -Be 'A1:B4'
        }
    }

    It 'keeps append freeze panes anchored to the existing table header' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelAppendFreeze.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )
        $moreRows = @(
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 150 }
        )

        $rows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales' -Title 'Sales Export' -FreezeTopRow
        $moreRows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -Append -TableName 'Sales' -FreezeTopRow

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $pane = $sheetXml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='sheetViews']/*[local-name()='sheetView']/*[local-name()='pane']")

        $pane.GetAttribute('ySplit') | Should -Be '2'
    }

    It 'supports autofit and validation list helpers' {
        $path = Join-Path $TestDrive 'DslExcelExtras.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Status = 'New' }
            [PSCustomObject]@{ Name = 'Beta'; Status = 'Done' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -Data $rows -TableName 'Items' -AutoFit
                Add-OfficeExcelValidationList -Range 'C2:C3' -Values 'New','In Progress','Done'
                Invoke-OfficeExcelAutoFit -Columns
            }
        }

        Test-Path $path | Should -BeTrue
    }

    It 'supports row/column helpers and reader metadata' {
        $path = Join-Path $TestDrive 'DslExcelReaders.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Value = 10 }
            [PSCustomObject]@{ Name = 'Beta'; Value = 20 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Value'
                Set-OfficeExcelColumn -Column 1 -StartRow 2 -Values 'Alpha', 'Beta'
                Set-OfficeExcelColumn -Column 2 -StartRow 2 -Values 10, 20
                Set-OfficeExcelNamedRange -Name 'ManualRange' -Range 'A1:B3'
                Add-OfficeExcelTable -Data $rows -TableName 'Sales' -StartRow 5
            }
        } | Out-Null

        $named = Get-OfficeExcelNamedRange -Path $path -Sheet 'Data' | Where-Object Name -eq 'ManualRange'
        $named | Should -Not -BeNullOrEmpty

        $tables = Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales'
        $tables | Should -Not -BeNullOrEmpty

        $doc = Get-OfficeExcel -Path $path
        try {
            $doc | Save-OfficeExcel | Out-Null
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'sets named ranges, formulas, and header/footer' {
        $path = Join-Path $TestDrive 'DslExcelMeta.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 10
                Set-OfficeExcelCell -Address 'B1' -Value 20
                Set-OfficeExcelFormula -Address 'C1' -Formula 'SUM(A1:B1)'
                Set-OfficeExcelNamedRange -Name 'Totals' -Range 'A1:C1'
                Set-OfficeExcelHeaderFooter -HeaderCenter 'Demo' -FooterRight 'Page &P of &N'
            }
        }

        Test-Path $path | Should -BeTrue
    }

    It 'supports advanced Excel data helpers' {
        $path = Join-Path $TestDrive 'DslExcelAdvancedData.xlsx'
        $rows = @(
            [PSCustomObject]@{
                Region = 'NA'
                Sales = 100
                Rate = 0.2
                CloseDate = [datetime]'2024-01-15'
                StartTime = [TimeSpan]'08:30:00'
                Note = 'OK'
            }
            [PSCustomObject]@{
                Region = 'EMEA'
                Sales = 200
                Rate = 0.45
                CloseDate = [datetime]'2024-02-20'
                StartTime = [TimeSpan]'09:15:00'
                Note = 'Check'
            }
            [PSCustomObject]@{
                Region = 'APAC'
                Sales = 150
                Rate = 0.33
                CloseDate = [datetime]'2024-03-10'
                StartTime = [TimeSpan]'10:05:00'
                Note = 'Review'
            }
        )

        $imagePath = New-TestOfficeImageFile -Directory $TestDrive

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -Data $rows -TableName 'Sales' -AutoFit
                Add-OfficeExcelAutoFilter -Range 'A1:F4'
                Invoke-OfficeExcelSort -Header 'Region'
                Set-OfficeExcelFreeze -TopRows 1
                Add-OfficeExcelConditionalRule -Range 'B2:B4' -Operator GreaterThan -Formula1 '150'
                Add-OfficeExcelConditionalDataBar -Range 'B2:B4' -Color '#92D050'
                Add-OfficeExcelConditionalColorScale -Range 'C2:C4' -StartColor '#FEE599' -EndColor '#6AA84F'
                Add-OfficeExcelConditionalIconSet -Range 'C2:C4'
                Add-OfficeExcelChart -TableName 'Sales' -Row 6 -Column 1 -Type ColumnClustered -Title 'Sales'
                Add-OfficeExcelImage -Address 'I1' -Path $imagePath -WidthPixels 64 -HeightPixels 64
                Set-OfficeExcelHyperlink -Address 'A2' -Url 'https://example.org' -Display 'Example'
                Add-OfficeExcelComment -Address 'B2' -Text 'Check sales'
                Add-OfficeExcelSparkline -DataRange 'B2:B4' -LocationRange 'H2:H4' -Type Column
            }
        }

        Test-Path $path | Should -BeTrue

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Count | Should -Be 1
            $sheet = $doc.Sheets[0]
            $sheet.Name | Should -Be 'Data'
            $sheet.HasComment(2, 2) | Should -BeTrue

            $cellText = $null
            $sheet.TryGetCellText(2, 1, [ref] $cellText) | Should -BeTrue
            $cellText | Should -Be 'Example'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'supports advanced Excel pivot, validation, and protection helpers' {
        $path = Join-Path $TestDrive 'DslExcelAdvancedPivot.xlsx'
        $rows = @(
            [PSCustomObject]@{
                Region = 'NA'
                Sales = 100
                Rate = 0.2
                CloseDate = [datetime]'2024-01-15'
                StartTime = [TimeSpan]'08:30:00'
                Note = 'OK'
            }
            [PSCustomObject]@{
                Region = 'EMEA'
                Sales = 200
                Rate = 0.45
                CloseDate = [datetime]'2024-02-20'
                StartTime = [TimeSpan]'09:15:00'
                Note = 'Check'
            }
            [PSCustomObject]@{
                Region = 'APAC'
                Sales = 150
                Rate = 0.33
                CloseDate = [datetime]'2024-03-10'
                StartTime = [TimeSpan]'10:05:00'
                Note = 'Review'
            }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -Data $rows -TableName 'Sales' -AutoFit
                Add-OfficeExcelPivotTable -SourceRange 'A1:F4' -DestinationCell 'J1' -RowField 'Region' -DataField 'Sales'
                Add-OfficeExcelValidationWholeNumber -Range 'B2:B4' -Operator Between -Formula1 1 -Formula2 1000 -AllowBlank:$false
                Add-OfficeExcelValidationDecimal -Range 'C2:C4' -Operator Between -Formula1 0.0 -Formula2 1.0
                Add-OfficeExcelValidationDate -Range 'D2:D4' -Operator GreaterThan -Formula1 ([datetime]'2024-01-01')
                Add-OfficeExcelValidationTime -Range 'E2:E4' -Operator GreaterThan -Formula1 ([TimeSpan]'08:00:00')
                Add-OfficeExcelValidationTextLength -Range 'F2:F4' -Operator Between -Formula1 1 -Formula2 20
                Add-OfficeExcelValidationCustomFormula -Range 'G2:G4' -Formula 'LEN(A2)>0'
                Protect-OfficeExcelSheet
                Unprotect-OfficeExcelSheet
                Protect-OfficeExcelSheet
            }
        }

        Test-Path $path | Should -BeTrue

        $pivotTables = @(Get-OfficeExcelPivotTable -Path $path -Name 'PivotTable')
        $pivotTables.Count | Should -Be 1

        $pivot = $pivotTables[0]
        $pivot.SourceRange | Should -Be 'A1:F4'
        $pivot.Location | Should -Match '^J1:[A-Z]+\d+$'
        $pivot.RowFields | Should -Contain 'Region'
        @($pivot.DataFields).Count | Should -BeGreaterThan 0
        $pivot.DataFields[0].FieldName | Should -Be 'Sales'

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Count | Should -Be 1
            $doc.Sheets[0].IsProtected | Should -BeTrue
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'supports advanced Excel page setup and visibility helpers' {
        $path = Join-Path $TestDrive 'DslExcelAdvancedLayout.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -Data $rows -TableName 'Sales' -AutoFit
                Set-OfficeExcelPageSetup -FitToWidth 1 -FitToHeight 0
                Set-OfficeExcelMargins -Preset Narrow
                Set-OfficeExcelOrientation -Orientation Landscape
                Set-OfficeExcelGridlines -Hide
                Set-OfficeExcelSheetVisibility -Hide
            }
        }

        Test-Path $path | Should -BeTrue

        $workbookXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/workbook.xml'
        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'

        $workbookSheet = $workbookXml.SelectSingleNode("/*[local-name()='workbook']/*[local-name()='sheets']/*[local-name()='sheet']")
        $workbookSheet.GetAttribute('name') | Should -Be 'Data'
        $workbookSheet.GetAttribute('state') | Should -Be 'hidden'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.HiddenSheetCount | Should -Be 1
        $summary.Sheets[0].State | Should -Be 'Hidden'

        $pageSetup = $sheetXml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='pageSetup']")
        $pageSetup.GetAttribute('fitToWidth') | Should -Be '1'
        $pageSetup.GetAttribute('fitToHeight') | Should -Be '0'
        $pageSetup.GetAttribute('orientation') | Should -Be 'landscape'

        $pageMargins = $sheetXml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='pageMargins']")
        $pageMargins.GetAttribute('left') | Should -Be '0.25'
        $pageMargins.GetAttribute('right') | Should -Be '0.25'
        $pageMargins.GetAttribute('top') | Should -Be '0.5'
        $pageMargins.GetAttribute('bottom') | Should -Be '0.5'
    }

    It 'counts threaded comments in workbook summaries' {
        $path = Join-Path $TestDrive 'DslExcelThreadedComments.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -Data $rows -TableName 'Sales'
            }
        }

        $document = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($path, $true)
        try {
            $worksheetPart = @($document.WorkbookPart.WorksheetParts)[0]
            $addPartMethod = [DocumentFormat.OpenXml.Packaging.WorksheetPart].GetMethods() |
                Where-Object { $_.Name -eq 'AddNewPart' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 0 } |
                Select-Object -First 1
            $threadedPart = $addPartMethod.MakeGenericMethod([DocumentFormat.OpenXml.Packaging.WorksheetThreadedCommentsPart]).Invoke($worksheetPart, @())
            $threadedComments = [DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments.ThreadedComments]::new()
            $threadedComment = [DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments.ThreadedComment]::new()
            $threadedComment.Ref = 'A2'
            $threadedComment.PersonId = '{00000000-0000-0000-0000-000000000001}'
            $threadedComment.Id = '{00000000-0000-0000-0000-000000000002}'
            $threadedComment.AppendChild([DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments.ThreadedCommentText]::new('Modern note')) | Out-Null
            $threadedComments.AppendChild($threadedComment) | Out-Null
            $threadedComments.Save($threadedPart)
        } finally {
            $document.Dispose()
        }

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.CommentCount | Should -Be 1
        $summary.Sheets[0].CommentCount | Should -Be 1
    }

    It 'adds a table of contents and reads ranges with the new Excel readers' {
        $path = Join-Path $TestDrive 'DslExcelNavigation.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -Data $rows -TableName 'Sales' -AutoFit
                Set-OfficeExcelNamedRange -Name 'SalesData' -Range 'A1:B3'
            }
            Add-OfficeExcelSheet -Name 'Notes' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Label', 'Value'
                Set-OfficeExcelRow -Row 2 -Values 'Generated', 'Yes'
            }
        } | Out-Null

        $usedRange = Get-OfficeExcelUsedRange -Path $path -Sheet 'Data' -AsDataTable
        $usedRange.Rows.Count | Should -Be 2
        $usedRange.Columns[0].ColumnName | Should -Be 'Region'
        $usedRange.Rows[0]['Region'] | Should -Be 'NA'

        Add-OfficeExcelTableOfContents -Path $path -IncludeNamedRanges -AddBackLinks

        $tocRows = @(Get-OfficeExcelRange -Path $path -Sheet 'TOC' -Range 'A3:C5' -AsHashtable)
        $tocRows.Count | Should -Be 2
        $tocRows[0]['Sheet'] | Should -Be 'Data'
        $tocRows[0]['Named Ranges'] | Should -Match 'SalesData'
        $tocRows[1]['Sheet'] | Should -Be 'Notes'

        $noteRows = @(Get-OfficeExcelRange -Path $path -Sheet 'Notes' -Range 'A1:B2')
        $noteRows.Count | Should -Be 1
        $noteRows[0].Label | Should -Be 'Generated'
        $noteRows[0].Value | Should -Be 'Yes'

        $dataRows = @(Get-OfficeExcelRange -Path $path -Sheet 'Data' -Range 'A1:B3')
        $dataRows.Count | Should -Be 2
        $dataRows[0].Region | Should -Be 'NA'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.SheetCount | Should -Be 3
        $summary.VisibleSheetCount | Should -Be 3
        $summary.TableCount | Should -Be 2
        $summary.NamedRangeCount | Should -Be 1
        $summary.HyperlinkCount | Should -BeGreaterThan 0
        $summary.Sheets.Name | Should -Contain 'Data'
        ($summary.Sheets | Where-Object Name -eq 'Data').UsedRange | Should -Be 'A1:B5'
        ($summary.Sheets | Where-Object Name -eq 'Data').Tables.Name | Should -Contain 'Sales'

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets[0].Name | Should -Be 'TOC'

            $backLink = $null
            $doc['Data'].TryGetCellText(5, 1, [ref] $backLink) | Should -BeTrue
            $backLink | Should -Be '← TOC'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'includes chartsheet charts in workbook summaries' {
        $path = Join-Path $TestDrive 'WorkbookWithChartSheet.xlsx'
        $archive = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            function Add-ZipTextEntry {
                param(
                    [Parameter(Mandatory)]
                    [System.IO.Compression.ZipArchive] $Archive,

                    [Parameter(Mandatory)]
                    [string] $EntryName,

                    [Parameter(Mandatory)]
                    [string] $Content
                )

                $entry = $Archive.CreateEntry($EntryName)
                $stream = $entry.Open()
                try {
                    $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
                    try {
                        $writer.Write($Content)
                    } finally {
                        $writer.Dispose()
                    }
                } finally {
                    $stream.Dispose()
                }
            }

            Add-ZipTextEntry -Archive $archive -EntryName '[Content_Types].xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>
'@
            Add-ZipTextEntry -Archive $archive -EntryName '_rels/.rels' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/workbook.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Revenue Chart" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/_rels/workbook.xml.rels' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/>
</Relationships>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/worksheets/sheet1.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1" t="str"><v>Region</v></c><c r="B1" t="str"><v>Revenue</v></c></row>
    <row r="2"><c r="A2" t="str"><v>EMEA</v></c><c r="B2"><v>42</v></c></row>
  </sheetData>
</worksheet>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/chartsheets/sheet1.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <drawing r:id="rId1"/>
</chartsheet>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/chartsheets/_rels/sheet1.xml.rels' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/drawings/drawing1.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="0" y="0"/><xdr:ext cx="6000000" cy="4000000"/>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/drawings/_rels/drawing1.xml.rels' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/charts/chart1.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea><c:layout/></c:plotArea></c:chart>
</c:chartSpace>
'@
        } finally {
            $archive.Dispose()
        }

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.SheetCount | Should -Be 2
        $summary.ChartCount | Should -Be 1
        ($summary.Sheets | Where-Object Name -eq 'Revenue Chart').ChartCount | Should -Be 1
    }

    It 'formats Excel charts with legend, labels, and style presets' {
        $path = Join-Path $TestDrive 'DslExcelChartFormatting.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -Data $rows -TableName 'Sales' -AutoFit
                $chart = Add-OfficeExcelChart -TableName 'Sales' -Row 6 -Column 1 -Type Pie -Title 'Revenue Mix' -PassThru
                $formattedChart = $chart |
                    Set-OfficeExcelChartLegend -Position Right |
                    Set-OfficeExcelChartDataLabels -ShowValue $true -ShowPercent $true -Position OutsideEnd -NumberFormat '0.0%' -SourceLinked:$false |
                    Set-OfficeExcelChartStyle -StyleId 251 -ColorStyleId 10

                $formattedChart | Should -Not -BeNullOrEmpty
            }
        } | Out-Null

        $entries = Get-ZipEntriesLocal -Path $path
        ($entries | Where-Object { $_ -like 'xl/drawings/charts/style*.xml' }).Count | Should -BeGreaterThan 0
        ($entries | Where-Object { $_ -like 'xl/drawings/charts/colors*.xml' }).Count | Should -BeGreaterThan 0

        $chartXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/drawings/charts/chart1.xml'
        $legendPosition = $chartXml.SelectSingleNode("/*[local-name()='chartSpace']/*[local-name()='chart']/*[local-name()='legend']/*[local-name()='legendPos']")
        $legendPosition | Should -Not -BeNullOrEmpty
        $legendPosition.GetAttribute('val') | Should -Be 'r'

        $dataLabels = $chartXml.SelectSingleNode("//*[local-name()='dLbls']")
        $dataLabels | Should -Not -BeNullOrEmpty
        $dataLabels.SelectSingleNode("*[local-name()='showVal']").GetAttribute('val') | Should -Be '1'
        $dataLabels.SelectSingleNode("*[local-name()='showPercent']").GetAttribute('val') | Should -Be '1'
        $dataLabels.SelectSingleNode("*[local-name()='dLblPos']").GetAttribute('val') | Should -Be 'outEnd'

        $numberFormat = $dataLabels.SelectSingleNode("*[local-name()='numFmt']")
        $numberFormat | Should -Not -BeNullOrEmpty
        $numberFormat.GetAttribute('formatCode') | Should -Be '0.0%'
    }

    It 'supports url images and smart hyperlink helpers' {
        $path = Join-Path $TestDrive 'DslExcelLinksAndImages.xlsx'
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive
        $imageUrl = [System.Uri]::new($imagePath).AbsoluteUri

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Reference'
                Set-OfficeExcelCell -Address 'B1' -Value 'Host'
                Set-OfficeExcelSmartHyperlink -Address 'A2' -Url 'https://datatracker.ietf.org/doc/html/rfc7208'
                Set-OfficeExcelHostHyperlink -Address 'B2' -Url 'https://learn.microsoft.com/office/open-xml/'
                Add-OfficeExcelImageFromUrl -Address 'D2' -Url $imageUrl -WidthPixels 32 -HeightPixels 32
                Add-OfficeExcelImage -Address 'E2' -Url $imageUrl -WidthPixels 24 -HeightPixels 24
            }
        } | Out-Null

        $entries = Get-ZipEntriesLocal -Path $path
        ($entries | Where-Object { $_ -like 'xl/media/*' }).Count | Should -BeGreaterThan 0

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $smartText = $null
            $hostText = $null
            $doc['Data'].TryGetCellText(2, 1, [ref] $smartText) | Should -BeTrue
            $doc['Data'].TryGetCellText(2, 2, [ref] $hostText) | Should -BeTrue
            $smartText | Should -Be 'RFC 7208'
            $hostText | Should -Be 'learn.microsoft.com'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $hyperlinks = $sheetXml.SelectNodes("/*[local-name()='worksheet']/*[local-name()='hyperlinks']/*[local-name()='hyperlink']")
        $hyperlinks.Count | Should -Be 2
    }

    It 'supports internal link helpers for summary sheets' {
        $path = Join-Path $TestDrive 'DslExcelInternalLinks.xlsx'
        $rows = @(
            [PSCustomObject]@{ Sheet = 'Alpha'; Target = 'Alpha' }
            [PSCustomObject]@{ Sheet = 'Beta'; Target = 'Beta' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Summary' -Content {
                Add-OfficeExcelTable -Data $rows -TableName 'SummaryTable' -AutoFit
                Set-OfficeExcelCell -Address 'D1' -Value 'Sheet'
                Set-OfficeExcelCell -Address 'D2' -Value 'Alpha'
                Set-OfficeExcelCell -Address 'D3' -Value 'Beta'
                Set-OfficeExcelInternalLinks -Range 'D2:D3'
                Set-OfficeExcelInternalLinksByHeader -Header 'Sheet' -TableName 'SummaryTable' -DisplayScript { param($text) "Open $text" }
            }
            Add-OfficeExcelSheet -Name 'Alpha' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Alpha Home'
            }
            Add-OfficeExcelSheet -Name 'Beta' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Beta Home'
            }
        } | Out-Null

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $summarySheet = $doc['Summary']
            $tableLink1 = $null
            $tableLink2 = $null
            $rangeLink1 = $null
            $rangeLink2 = $null
            $summarySheet.TryGetCellText(2, 1, [ref] $tableLink1) | Should -BeTrue
            $summarySheet.TryGetCellText(3, 1, [ref] $tableLink2) | Should -BeTrue
            $summarySheet.TryGetCellText(2, 4, [ref] $rangeLink1) | Should -BeTrue
            $summarySheet.TryGetCellText(3, 4, [ref] $rangeLink2) | Should -BeTrue
            $tableLink1 | Should -Be 'Open Alpha'
            $tableLink2 | Should -Be 'Open Beta'
            $rangeLink1 | Should -Be 'Alpha'
            $rangeLink2 | Should -Be 'Beta'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $hyperlinks = $sheetXml.SelectNodes("/*[local-name()='worksheet']/*[local-name()='hyperlinks']/*[local-name()='hyperlink']")
        $hyperlinks.Count | Should -Be 4
    }

    It 'supports external URL link helpers for summary sheets' {
        $path = Join-Path $TestDrive 'DslExcelUrlLinks.xlsx'
        $rows = @(
            [PSCustomObject]@{ RFC = 'rfc7208'; Spec = 'rfc5321' }
            [PSCustomObject]@{ RFC = 'rfc7489'; Spec = 'rfc1035' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Summary' -Content {
                Add-OfficeExcelTable -Data $rows -TableName 'LinksTable' -AutoFit
                Set-OfficeExcelCell -Address 'D1' -Value 'Spec'
                Set-OfficeExcelCell -Address 'D2' -Value 'rfc5321'
                Set-OfficeExcelCell -Address 'D3' -Value 'rfc1035'

                Set-OfficeExcelUrlLinksByHeader -Header 'RFC' -TableName 'LinksTable' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" } -TitleScript { param($text) "Open $text" }
                Set-OfficeExcelUrlLinks -Range 'D2:D3' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" }
            }
        } | Out-Null

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $summarySheet = $doc['Summary']
            $tableLink1 = $null
            $tableLink2 = $null
            $rangeLink1 = $null
            $rangeLink2 = $null
            $summarySheet.TryGetCellText(2, 1, [ref] $tableLink1) | Should -BeTrue
            $summarySheet.TryGetCellText(3, 1, [ref] $tableLink2) | Should -BeTrue
            $summarySheet.TryGetCellText(2, 4, [ref] $rangeLink1) | Should -BeTrue
            $summarySheet.TryGetCellText(3, 4, [ref] $rangeLink2) | Should -BeTrue
            $tableLink1 | Should -Be 'Open rfc7208'
            $tableLink2 | Should -Be 'Open rfc7489'
            $rangeLink1 | Should -Be 'RFC 5321'
            $rangeLink2 | Should -Be 'RFC 1035'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $hyperlinks = $sheetXml.SelectNodes("/*[local-name()='worksheet']/*[local-name()='hyperlinks']/*[local-name()='hyperlink']")
        $hyperlinks.Count | Should -Be 4
    }
}
