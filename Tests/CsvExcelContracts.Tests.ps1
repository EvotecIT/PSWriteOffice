BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        $sourceRoot = Join-Path (Join-Path (Join-Path $PSScriptRoot '..') 'Sources') 'PSWriteOffice'

        if (-not $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES) {
            $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
        }

        if (-not $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION) {
            $releasePath = Join-Path (Join-Path $sourceRoot 'bin') 'Release'
            $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = if (Test-Path $releasePath) { 'Release' } else { 'Debug' }
        }

        Join-Path (Join-Path $PSScriptRoot '..') 'PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'CSV and Excel mutation contracts' {
    It 'does not create a CSV file when ConvertTo-OfficeCsv is invoked with WhatIf' {
        $path = Join-Path $TestDrive 'whatif.csv'

        1..3 |
            ForEach-Object { [pscustomobject]@{ Name = "Row$_"; Value = $_ } } |
            ConvertTo-OfficeCsv -OutputPath $path -WhatIf

        Test-Path -LiteralPath $path | Should -BeFalse
    }

    It 'imports a CSV source into a new Excel workbook' {
        $csv = Join-Path $TestDrive 'source.csv'
        $xlsx = Join-Path $TestDrive 'created.xlsx'
        1..3 |
            ForEach-Object { [pscustomobject]@{ Name = "Row$_"; Value = $_ } } |
            Export-Csv -Path $csv -NoTypeInformation

        $result = Import-OfficeExcelDelimitedText -Path $xlsx -SourcePath $csv -PassThru

        Test-Path -LiteralPath $xlsx | Should -BeTrue
        $result.RowCount | Should -Be 3
        $rows = Import-OfficeExcel -Path $xlsx -WorksheetName Import
        $rows.Count | Should -Be 3
        $rows[0].Name | Should -Be 'Row1'
    }

    It 'exports PowerShell objects using the first row schema and blanks missing later values' {
        $xlsx = Join-Path $TestDrive 'projected.xlsx'

        @(
            [pscustomobject]@{ Name = 'Row1'; Value = 1 }
            [pscustomobject]@{ Name = 'Row2'; Extra = 'Ignored' }
        ) | Export-OfficeExcel -Path $xlsx -WorksheetName Data -TableName Data

        $rows = @(Import-OfficeExcel -Path $xlsx -WorksheetName Data)
        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be 'Row1'
        $rows[0].Value | Should -Be 1
        $rows[1].Name | Should -Be 'Row2'
        $rows[1].Value | Should -Be ''
        $rows[1].PSObject.Properties.Name | Should -Not -Contain 'Extra'
    }

    It 'does not append Excel table rows when Add-OfficeExcelTableRow is invoked with WhatIf' {
        $xlsx = Join-Path $TestDrive 'table.xlsx'
        [pscustomobject]@{ Name = 'Row1'; Value = 1 } |
            Export-OfficeExcel -Path $xlsx -WorksheetName Data -TableName Data

        [pscustomobject]@{ Name = 'Row2'; Value = 2 } |
            Add-OfficeExcelTableRow -Path $xlsx -Sheet Data -TableName Data -WhatIf

        $rows = Import-OfficeExcel -Path $xlsx -WorksheetName Data
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'Row1'
    }
}
