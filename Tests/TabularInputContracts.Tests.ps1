BeforeAll {
    if (-not $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES) {
        $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
    }
    if (-not $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION) {
        $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = 'Debug'
    }

    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -Force -ErrorAction Stop

    if (-not ('PSWriteOffice.Tests.TabularContractRow' -as [type])) {
        Add-Type -TypeDefinition @'
using System.Collections.Generic;

namespace PSWriteOffice.Tests {
    public sealed class TabularContractRow {
        public string Name { get; set; }
        public string[] Tags { get; set; }
        public Dictionary<string, object> Metadata { get; set; }
    }

    public sealed class ReadOnlyTabularContractRow : IReadOnlyDictionary<string, object> {
        private readonly KeyValuePair<string, object>[] _entries;
        private readonly Dictionary<string, object> _lookup;

        public ReadOnlyTabularContractRow(IDictionary<string, object> values) {
            _lookup = new Dictionary<string, object>(values, System.StringComparer.OrdinalIgnoreCase);
            _entries = new List<KeyValuePair<string, object>>(values).ToArray();
        }

        public object this[string key] { get { return _lookup[key]; } }
        public IEnumerable<string> Keys { get { foreach (var entry in _entries) yield return entry.Key; } }
        public IEnumerable<object> Values { get { foreach (var entry in _entries) yield return entry.Value; } }
        public int Count { get { return _entries.Length; } }
        public bool ContainsKey(string key) { return _lookup.ContainsKey(key); }
        public bool TryGetValue(string key, out object value) { return _lookup.TryGetValue(key, out value); }
        public IEnumerator<KeyValuePair<string, object>> GetEnumerator() {
            return ((IEnumerable<KeyValuePair<string, object>>)_entries).GetEnumerator();
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { return _entries.GetEnumerator(); }
    }
}
'@
    }

    function New-TestTabularContractReader {
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Tags', [object])
        [void] $table.Columns.Add('Metadata', [object])
        [void] $table.Rows.Add(
            'Alpha',
            [object] @('One', 'Two'),
            [object] ([ordered]@{ Region = 'EU'; Tier = 1 }))
        return ,$table.CreateDataReader()
    }
}

Describe 'Shared tabular input contracts' {
    It 'exposes nested value formatting consistently across table outputs' {
        $commands = @(
            'Add-OfficeExcelTable'
            'Export-OfficeExcel'
            'Add-OfficeExcelReportTable'
            'Add-OfficeWordTable'
            'Add-OfficePowerPointTable'
            'Add-OfficePdfTable'
            'Add-OfficeMarkdownTable'
            'ConvertTo-OfficeMarkdown'
        )

        foreach ($commandName in $commands) {
            $parameters = (Get-Command $commandName -ErrorAction Stop).Parameters.Keys
            $parameters | Should -Contain 'CollectionSeparator'
            $parameters | Should -Contain 'DictionaryEntrySeparator'
            $parameters | Should -Contain 'DictionaryKeyValueSeparator'
        }
    }

    It 'renders object and dictionary row families through Excel report tables' {
        $genericDictionary = [System.Collections.Generic.Dictionary[string, object]]::new()
        $genericDictionary.Add('Name', 'GenericDictionary')
        $genericDictionary.Add('Value', 3)

        $cases = @(
            [pscustomobject]@{
                Name  = 'PSCustomObject'
                Input = [pscustomobject]@{ Name = 'PSCustomObject'; Value = 1 }
            }
            [pscustomobject]@{
                Name  = 'OrderedDictionary'
                Input = [ordered]@{ Name = 'OrderedDictionary'; Value = 2 }
            }
            [pscustomobject]@{
                Name  = 'GenericDictionary'
                Input = $genericDictionary
            }
            [pscustomobject]@{
                Name  = 'ClrObject'
                Input = [System.Collections.DictionaryEntry]::new('ClrObject', 4)
            }
        )

        foreach ($case in $cases) {
            $path = Join-Path $TestDrive ("Report-{0}.xlsx" -f $case.Name)
            $inputRow = $case.Input

            New-OfficeExcel -Path $path {
                Add-OfficeExcelReportSheet -Name 'Report' {
                    Add-OfficeExcelReportTable -InputObject $inputRow
                }
            }

            $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Report')
            $rows.Count | Should -Be 1

            if ($case.Name -eq 'ClrObject') {
                $rows[0].Key | Should -Be 'ClrObject'
                $rows[0].Value | Should -Be 4
            } else {
                $rows[0].Name | Should -Be $case.Name
                $rows[0].Value | Should -BeGreaterThan 0
            }
        }
    }

    It 'unions columns found only in later heterogeneous PowerShell rows' {
        $path = Join-Path $TestDrive 'Report-HeterogeneousRows.xlsx'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha' }
            [pscustomobject]@{ Name = 'Beta'; Department = 'Operations' }
            [ordered]@{ Name = 'Gamma'; Score = 7 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Report' {
                Add-OfficeExcelReportTable -InputObject $rows
            }
        }

        $result = @(Import-OfficeExcel -Path $path -WorksheetName 'Report')
        $result | Should -HaveCount 3
        $result[0].PSObject.Properties.Name | Should -Contain 'Department'
        $result[0].PSObject.Properties.Name | Should -Contain 'Score'
        $result[0].Department | Should -BeNullOrEmpty
        $result[1].Department | Should -Be 'Operations'
        $result[2].Score | Should -Be 7
    }

    It 'renders ADO.NET tabular input families without DataRow metadata' {
        $createTable = {
            $table = [System.Data.DataTable]::new('Rows')
            [void] $table.Columns.Add('Name', [string])
            [void] $table.Columns.Add('Value', [int])
            [void] $table.Rows.Add('Alpha', 7)
            return ,$table
        }

        $inputFactories = [ordered]@{
            DataTable   = { & $createTable }
            DataView    = { (& $createTable).DefaultView }
            DataRow     = { (& $createTable).Rows[0] }
            IDataReader = { return ,((& $createTable).CreateDataReader()) }
        }

        foreach ($entry in $inputFactories.GetEnumerator()) {
            $path = Join-Path $TestDrive ("Report-{0}.xlsx" -f $entry.Key)
            $inputRows = & $entry.Value
            try {
                New-OfficeExcel -Path $path {
                    Add-OfficeExcelReportSheet -Name 'Report' {
                        Add-OfficeExcelReportTable -InputObject $inputRows
                    }
                }
            } finally {
                if ($inputRows -is [System.Data.IDataReader]) {
                    $inputRows.Dispose()
                }
            }

            $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Report')
            $rows.Count | Should -Be 1 -Because $entry.Key
            $rows[0].Name | Should -Be 'Alpha' -Because $entry.Key
            $rows[0].Value | Should -Be 7 -Because $entry.Key
            $rows[0].PSObject.Properties.Name | Should -Not -Contain 'RowError'
            $rows[0].PSObject.Properties.Name | Should -Not -Contain 'Table'
        }
    }

    It 'renders scalar input as a single Value column' {
        $path = Join-Path $TestDrive 'Report-Scalar.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Report' {
                Add-OfficeExcelReportTable -InputObject 'Scalar value'
            }
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Report')
        $rows.Count | Should -Be 1
        $rows[0].Value | Should -Be 'Scalar value'
    }

    It 'formats nested collections and dictionaries identically in Excel and Markdown tables' {
        $inputRow = [pscustomobject]@{
            Name     = 'Alpha'
            Tags     = @('One', 'Two')
            Metadata = [ordered]@{ Region = 'EU'; Tier = 1 }
        }
        $excelPath = Join-Path $TestDrive 'Report-NestedValues.xlsx'

        New-OfficeExcel -Path $excelPath {
            Add-OfficeExcelReportSheet -Name 'Report' {
                Add-OfficeExcelReportTable -InputObject $inputRow `
                    -CollectionSeparator ' | ' `
                    -DictionaryEntrySeparator '; ' `
                    -DictionaryKeyValueSeparator ': '
            }
        }

        $excelRows = @(Import-OfficeExcel -Path $excelPath -WorksheetName 'Report')
        $excelRows[0].Tags | Should -Be 'One | Two'
        $excelRows[0].Metadata | Should -Be 'Region: EU; Tier: 1'

        $markdownPath = Join-Path $TestDrive 'Report-NestedValues.md'
        New-OfficeMarkdown -Path $markdownPath {
            Add-OfficeMarkdownTable -InputObject $inputRow `
                -CollectionSeparator ' | ' `
                -DictionaryEntrySeparator '; ' `
                -DictionaryKeyValueSeparator ': '
        }

        $markdown = Get-Content -LiteralPath $markdownPath -Raw
        $markdown | Should -Match ([regex]::Escape('One \| Two'))
        $markdown | Should -Match 'Region: EU; Tier: 1'

        $convertedMarkdown = $inputRow | ConvertTo-OfficeMarkdown `
            -CollectionSeparator ' | ' `
            -DictionaryEntrySeparator '; ' `
            -DictionaryKeyValueSeparator ': '
        $convertedMarkdown | Should -Match ([regex]::Escape('One \| Two'))
        $convertedMarkdown | Should -Match 'Region: EU; Tier: 1'
    }

    It 'supports column inclusion and exclusion on Excel report tables' {
        $path = Join-Path $TestDrive 'Report-Projection.xlsx'
        $inputRow = [pscustomobject]@{ Name = 'Alpha'; Value = 1; Internal = 'Hidden' }

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Report' {
                Add-OfficeExcelReportTable -InputObject $inputRow -Property Value, Name -ExcludeProperty Internal
            }
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Report')
        $rows[0].PSObject.Properties.Name | Should -Be @('Value', 'Name')
        $rows[0].Value | Should -Be 1
        $rows[0].Name | Should -Be 'Alpha'
    }

    It 'streams IDataReader rows with shared nested formatting into direct Excel exports' {
        $path = Join-Path $TestDrive 'DirectReader.xlsx'
        $reader = New-TestTabularContractReader
        try {
            Export-OfficeExcel -Path $path -InputObject $reader `
                -CollectionSeparator ' | ' `
                -DictionaryEntrySeparator '; ' `
                -DictionaryKeyValueSeparator ': '
        } finally {
            $reader.Dispose()
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sheet1')
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'Alpha'
        $rows[0].Tags | Should -Be 'One | Two'
        $rows[0].Metadata | Should -Be 'Region: EU; Tier: 1'
    }

    It 'normalizes dictionary fast-path exports before OfficeIMO writes the workbook' {
        $path = Join-Path $TestDrive 'DirectDictionary.xlsx'
        $row = [ordered]@{
            Name = 'Alpha'
            Tags = @('One', 'Two')
            Metadata = [ordered]@{ Region = 'EU'; Tier = 1 }
        }

        Export-OfficeExcel -Path $path -InputObject $row `
            -CollectionSeparator ' | ' `
            -DictionaryEntrySeparator '; ' `
            -DictionaryKeyValueSeparator ': '

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sheet1')
        $rows[0].Tags | Should -Be 'One | Two'
        $rows[0].Metadata | Should -Be 'Region: EU; Tier: 1'
    }

    It 'materializes IDataReader rows consistently across Word PowerPoint PDF and Markdown tables' {
        $wordPath = Join-Path $TestDrive 'Reader.docx'
        $reader = New-TestTabularContractReader
        try {
            New-OfficeWord -Path $wordPath {
                Add-OfficeWordTable -InputObject $reader -Style TableGrid `
                    -CollectionSeparator ' | ' `
                    -DictionaryEntrySeparator '; ' `
                    -DictionaryKeyValueSeparator ': '
            } | Out-Null
        } finally {
            $reader.Dispose()
        }

        $word = Get-OfficeWord -Path $wordPath -ReadOnly
        try {
            $word.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Alpha'
            $word.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'One | Two'
            $word.Tables[0].Rows[1].Cells[2].Paragraphs[0].Text | Should -Be 'Region: EU; Tier: 1'
        } finally {
            $word.Dispose()
        }

        $powerPointPath = Join-Path $TestDrive 'Reader.pptx'
        $presentation = New-OfficePowerPoint -FilePath $powerPointPath -NoSave
        try {
            $slide = Add-OfficePowerPointSlide -Presentation $presentation
            $reader = New-TestTabularContractReader
            try {
                $table = Add-OfficePowerPointTable -Slide $slide -InputObject $reader `
                    -CollectionSeparator ' | ' `
                    -DictionaryEntrySeparator '; ' `
                    -DictionaryKeyValueSeparator ': '
            } finally {
                $reader.Dispose()
            }

            $table.GetCell(1, 0).Text | Should -Be 'Alpha'
            $table.GetCell(1, 1).Text | Should -Be 'One | Two'
            $table.GetCell(1, 2).Text | Should -Be 'Region: EU; Tier: 1'
        } finally {
            Close-OfficePowerPoint -Presentation $presentation
        }

        $pdfPath = Join-Path $TestDrive 'Reader.pdf'
        $reader = New-TestTabularContractReader
        try {
            New-OfficePdf -Path $pdfPath {
                Add-OfficePdfTable -InputObject $reader `
                    -CollectionSeparator ' | ' `
                    -DictionaryEntrySeparator '; ' `
                    -DictionaryKeyValueSeparator ': '
            } | Out-Null
        } finally {
            $reader.Dispose()
        }

        $pdfText = Get-OfficePdfText -Path $pdfPath
        $pdfText | Should -Match 'Alpha'
        $pdfText | Should -Match 'One \| Two'
        $pdfText | Should -Match 'Region: EU; Tier: 1'

        $markdownPath = Join-Path $TestDrive 'Reader.md'
        $reader = New-TestTabularContractReader
        try {
            New-OfficeMarkdown -Path $markdownPath {
                Add-OfficeMarkdownTable -InputObject $reader `
                    -CollectionSeparator ' | ' `
                    -DictionaryEntrySeparator '; ' `
                    -DictionaryKeyValueSeparator ': '
            }
        } finally {
            $reader.Dispose()
        }

        $markdown = Get-Content -LiteralPath $markdownPath -Raw
        $markdown | Should -Match 'Alpha'
        $markdown | Should -Match ([regex]::Escape('One \| Two'))
        $markdown | Should -Match 'Region: EU; Tier: 1'
    }

    It 'projects CLR rows before nested values reach format-specific writers' {
        $metadata = [System.Collections.Generic.Dictionary[string, object]]::new()
        $metadata.Add('Region', 'EU')
        $metadata.Add('Tier', 1)
        $row = [PSWriteOffice.Tests.TabularContractRow]::new()
        $row.Name = 'Alpha'
        $row.Tags = @('One', 'Two')
        $row.Metadata = $metadata

        $markdown = $row | ConvertTo-OfficeMarkdown `
            -CollectionSeparator ' | ' `
            -DictionaryEntrySeparator '; ' `
            -DictionaryKeyValueSeparator ': '

        $markdown | Should -Match 'Alpha'
        $markdown | Should -Match ([regex]::Escape('One \| Two'))
        $markdown | Should -Match 'Region: EU; Tier: 1'
    }

    It 'keeps IReadOnlyDictionary-only rows as named columns across table formats' {
        $values = [System.Collections.Generic.Dictionary[string, object]]::new()
        $values.Add('Name', 'Alpha')
        $values.Add('Tags', @('One', 'Two'))
        $values.Add('Metadata', [ordered]@{ Region = 'EU'; Tier = 1 })
        $row = [PSWriteOffice.Tests.ReadOnlyTabularContractRow]::new($values)

        $exportPath = Join-Path $TestDrive 'ReadOnlyDictionaryExport.xlsx'
        Export-OfficeExcel -Path $exportPath -InputObject $row -CollectionSeparator ' | '
        $exportRows = @(Import-OfficeExcel -Path $exportPath -WorksheetName 'Sheet1')
        $exportRows[0].Name | Should -Be 'Alpha'
        $exportRows[0].Tags | Should -Be 'One | Two'

        $tablePath = Join-Path $TestDrive 'ReadOnlyDictionaryTable.xlsx'
        New-OfficeExcel -Path $tablePath {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $row -TableName 'ReadOnlyRows' -CollectionSeparator ' | '
            }
        }
        $tableRows = @(Import-OfficeExcel -Path $tablePath -WorksheetName 'Data')
        $tableRows[0].Name | Should -Be 'Alpha'
        $tableRows[0].Tags | Should -Be 'One | Two'

        $reportPath = Join-Path $TestDrive 'ReadOnlyDictionaryReport.xlsx'
        New-OfficeExcel -Path $reportPath {
            Add-OfficeExcelReportSheet -Name 'Report' {
                Add-OfficeExcelReportTable -InputObject $row -CollectionSeparator ' | '
            }
        }
        $reportRows = @(Import-OfficeExcel -Path $reportPath -WorksheetName 'Report')
        $reportRows[0].Name | Should -Be 'Alpha'
        $reportRows[0].Tags | Should -Be 'One | Two'

        $markdown = $row | ConvertTo-OfficeMarkdown -CollectionSeparator ' | '
        $markdown | Should -Match 'Name'
        $markdown | Should -Match 'Alpha'
        $markdown | Should -Match ([regex]::Escape('One \| Two'))

        $wordPath = Join-Path $TestDrive 'ReadOnlyDictionary.docx'
        New-OfficeWord -Path $wordPath {
            Add-OfficeWordTable -InputObject $row -Style TableGrid -CollectionSeparator ' | '
        } | Out-Null
        $word = Get-OfficeWord -Path $wordPath -ReadOnly
        try {
            $word.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Name'
            $word.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Alpha'
            $word.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'One | Two'
        } finally {
            $word.Dispose()
        }

        $powerPointPath = Join-Path $TestDrive 'ReadOnlyDictionary.pptx'
        $presentation = New-OfficePowerPoint -FilePath $powerPointPath -NoSave
        try {
            $slide = Add-OfficePowerPointSlide -Presentation $presentation
            $table = Add-OfficePowerPointTable -Slide $slide -InputObject $row -CollectionSeparator ' | '
            $table.GetCell(0, 0).Text | Should -Be 'Name'
            $table.GetCell(1, 0).Text | Should -Be 'Alpha'
            $table.GetCell(1, 1).Text | Should -Be 'One | Two'
        } finally {
            Close-OfficePowerPoint -Presentation $presentation
        }

        $pdfPath = Join-Path $TestDrive 'ReadOnlyDictionary.pdf'
        New-OfficePdf -Path $pdfPath {
            Add-OfficePdfTable -InputObject $row -CollectionSeparator ' | '
        } | Out-Null
        $pdfText = Get-OfficePdfText -Path $pdfPath
        $pdfText | Should -Match 'Name'
        $pdfText | Should -Match 'Alpha'
        $pdfText | Should -Match 'One \| Two'
    }
}
