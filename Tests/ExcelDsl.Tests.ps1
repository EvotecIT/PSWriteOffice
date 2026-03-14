BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Force -Global

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

    It 'supports advanced Excel helpers' {
        function Get-ZipXmlDocumentLocal {
            param(
                [Parameter(Mandatory)]
                [string] $Path,

                [Parameter(Mandatory)]
                [string] $Entry
            )

            $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
            try {
                $zipEntry = $archive.GetEntry($Entry)
                if (-not $zipEntry) {
                    throw "Zip entry '$Entry' not found in '$Path'."
                }

                $stream = $zipEntry.Open()
                try {
                    $reader = [System.IO.StreamReader]::new($stream)
                    try {
                        return [xml] $reader.ReadToEnd()
                    } finally {
                        $reader.Dispose()
                    }
                } finally {
                    $stream.Dispose()
                }
            } finally {
                $archive.Dispose()
            }
        }

        $path = Join-Path $TestDrive 'DslExcelAdvanced.xlsx'
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
                Add-OfficeExcelPivotTable -SourceRange 'A1:F4' -DestinationCell 'J1' -RowField 'Region' -DataField 'Sales'
                Protect-OfficeExcelSheet
                Unprotect-OfficeExcelSheet
                Add-OfficeExcelValidationWholeNumber -Range 'B2:B4' -Operator Between -Formula1 1 -Formula2 1000 -AllowBlank:$false
                Add-OfficeExcelValidationDecimal -Range 'C2:C4' -Operator Between -Formula1 0.0 -Formula2 1.0
                Add-OfficeExcelValidationDate -Range 'D2:D4' -Operator GreaterThan -Formula1 ([datetime]'2024-01-01')
                Add-OfficeExcelValidationTime -Range 'E2:E4' -Operator GreaterThan -Formula1 ([TimeSpan]'08:00:00')
                Add-OfficeExcelValidationTextLength -Range 'F2:F4' -Operator Between -Formula1 1 -Formula2 20
                Add-OfficeExcelValidationCustomFormula -Range 'G2:G4' -Formula 'LEN(A2)>0'
                Set-OfficeExcelPageSetup -FitToWidth 1 -FitToHeight 0
                Set-OfficeExcelMargins -Preset Narrow
                Set-OfficeExcelOrientation -Orientation Landscape
                Set-OfficeExcelGridlines -Hide
                Set-OfficeExcelSheetVisibility -Hide
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
            $sheet = $doc.Sheets[0]
            $sheet.Name | Should -Be 'Data'
            $sheet.IsProtected | Should -BeTrue
            $sheet.HasComment(2, 2) | Should -BeTrue

            $cellText = $null
            $sheet.TryGetCellText(2, 1, [ref] $cellText) | Should -BeTrue
            $cellText | Should -Be 'Example'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $workbookXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/workbook.xml'
        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'

        $workbookSheet = $workbookXml.SelectSingleNode("/*[local-name()='workbook']/*[local-name()='sheets']/*[local-name()='sheet']")
        $workbookSheet.GetAttribute('name') | Should -Be 'Data'
        $workbookSheet.GetAttribute('state') | Should -Be 'hidden'

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
}
