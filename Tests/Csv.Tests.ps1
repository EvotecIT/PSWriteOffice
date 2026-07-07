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

Describe 'CSV cmdlets' {
    It 'exposes NoHeader instead of bool header toggles on CSV commands' {
        (Get-Command ConvertTo-OfficeCsv).Parameters.Keys | Should -Contain 'NoHeader'
        (Get-Command Export-OfficeCsv).Parameters.Keys | Should -Contain 'NoHeader'
        (Get-Command Get-OfficeCsv).Parameters.Keys | Should -Contain 'NoHeader'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'NoHeader'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'AsDataTable'

        (Get-Command ConvertTo-OfficeCsv).Parameters.Keys | Should -Not -Contain 'IncludeHeader'
        (Get-Command Export-OfficeCsv).Parameters.Keys | Should -Not -Contain 'IncludeHeader'
        (Get-Command Get-OfficeCsv).Parameters.Keys | Should -Not -Contain 'HasHeaderRow'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Not -Contain 'HasHeaderRow'
    }

    It 'exposes CSV row import with idiomatic command names' {
        (Get-Command Import-OfficeCsv).CommandType | Should -Be 'Cmdlet'
        (Get-Command ConvertFrom-OfficeCsv).CommandType | Should -Be 'Cmdlet'
        { Get-Command Get-OfficeCsvData -ErrorAction Stop } | Should -Throw
    }

    It 'converts objects to CSV and reads them back' {
        $rows = @(
            [pscustomobject]@{ Region = 'NA'; Revenue = 100 }
            [pscustomobject]@{ Region = 'EMEA'; Revenue = 200 }
        )

        $csvText = @($rows | ConvertTo-OfficeCsv)
        $csvText.Count | Should -Be 3
        $csvText[0] | Should -Be 'Region,Revenue'
        $csvText[1] | Should -Be 'NA,100'
        $csvText[2] | Should -Be 'EMEA,200'

        $path = Join-Path $TestDrive 'data.csv'
        $rows | Export-OfficeCsv -Path $path | Out-Null

        Test-Path $path | Should -BeTrue

        $data = Import-OfficeCsv -Path $path
        $data.Count | Should -Be 2
        $data[0].Region | Should -Be 'NA'
    }

    It 'imports CSV rows as a DataTable for database workflows' {
        $path = Join-Path $TestDrive 'datatable.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`nBeta,2" -Encoding UTF8

        $table = Import-OfficeCsv -Path $path -AsDataTable

        $table.GetType() | Should -Be ([System.Data.DataTable])
        $table.TableName | Should -Be 'datatable'
        $table.Columns.Count | Should -Be 2
        $table.Columns['Name'].DataType | Should -Be ([string])
        $table.Rows.Count | Should -Be 2
        $table.Rows[0]['Name'] | Should -Be 'Alpha'
        $table.Rows[1]['Value'] | Should -Be '2'
    }

    It 'preserves normalized empty DataTable fields' {
        $path = Join-Path $TestDrive 'datatable-missing.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`nBeta" -Encoding UTF8

        $table = Import-OfficeCsv -Path $path -AsDataTable

        $table.Rows[1]['Value'] | Should -Be ''
    }

    It 'rejects conflicting CSV table and hashtable output modes' {
        $path = Join-Path $TestDrive 'datatable-conflict.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -Encoding UTF8

        { Import-OfficeCsv -Path $path -AsDataTable -AsHashtable -ErrorAction Stop } |
            Should -Throw '*AsDataTable*AsHashtable*'
    }

    It 'writes to files using the Path alias' {
        $path = Join-Path $TestDrive 'path-alias.csv'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        Test-Path $path | Should -BeTrue
        Get-Content -Path $path -Raw | Should -Match 'Alpha'
    }

    It 'writes to literal file paths without wildcard expansion' {
        $path = Join-Path $TestDrive 'literal[export].csv'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -LiteralPath $path

        Test-Path -LiteralPath $path | Should -BeTrue
        (Import-OfficeCsv -LiteralPath $path)[0].Name | Should -Be 'Alpha'
    }

    It 'does not overwrite an existing CSV file when NoClobber is specified' {
        $path = Join-Path $TestDrive 'no-clobber.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nOriginal,1" -Encoding UTF8

        {
            [pscustomobject]@{ Name = 'New'; Value = 2 } |
                Export-OfficeCsv -Path $path -NoClobber -ErrorAction Stop
        } | Should -Throw

        (Import-OfficeCsv -Path $path)[0].Name | Should -Be 'Original'
    }

    It 'appends object rows using the existing CSV header order' {
        $path = Join-Path $TestDrive 'append-order.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        [pscustomobject]@{ Value = 2; Name = 'Beta'; Extra = 'Ignored' } |
            Export-OfficeCsv -Path $path -Append

        $raw = Get-Content -LiteralPath $path
        $data = Import-OfficeCsv -Path $path

        $raw | Should -Be @('Name,Value', 'Alpha,1', 'Beta,2')
        $data.Count | Should -Be 2
        $data[1].Name | Should -Be 'Beta'
        $data[1].Value | Should -Be '2'
    }

    It 'uses existing CSV header order when appending without writing headers' {
        $path = Join-Path $TestDrive 'append-order-no-header.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        [pscustomobject]@{ Value = 2; Name = 'Beta' } |
            Export-OfficeCsv -Path $path -Append -NoHeader

        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1', 'Beta,2')
    }

    It 'starts appended rows on a new record when the existing file has no trailing newline' {
        $path = Join-Path $TestDrive 'append-no-newline.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -NoNewline -Encoding UTF8

        [pscustomobject]@{ Name = 'Beta'; Value = 2 } |
            Export-OfficeCsv -Path $path -Append

        $raw = Get-Content -LiteralPath $path -Raw
        $raw | Should -Match "Alpha,1(`r`n|`n|`r)Beta,2"
        $data = @(Import-OfficeCsv -Path $path)
        $data.Count | Should -Be 2
        $data[1].Name | Should -Be 'Beta'
    }

    It 'treats whitespace-only append targets as new CSV files' {
        $path = Join-Path $TestDrive 'append-empty-target.csv'
        Set-Content -LiteralPath $path -Value "`r`n" -NoNewline -Encoding UTF8

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -Append

        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1')
    }

    It 'does not infer headers when appending to a headerless CSV' {
        $path = Join-Path $TestDrive 'append-headerless.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -NoHeader

        [pscustomobject]@{ Name = 'Beta'; Value = 2 } |
            Export-OfficeCsv -Path $path -Append -NoHeader

        Get-Content -LiteralPath $path | Should -Be @('Alpha,1', 'Beta,2')
    }

    It 'preserves BOM-detected CSV encoding when appending' {
        $path = Join-Path $TestDrive 'append-utf16.csv'
        $firstName = 'Za' + [char] 0x017C + [char] 0x00F3 + [char] 0x0142 + [char] 0x0107
        $secondName = [string] [char] 0x0141 + [char] 0x00F3 + 'd' + [char] 0x017A
        Set-Content -LiteralPath $path -Value "Name,Value`r`n$firstName,1" -NoNewline -Encoding Unicode

        [pscustomobject]@{ Name = $secondName; Value = 2 } |
            Export-OfficeCsv -Path $path -Append

        $bytes = [System.IO.File]::ReadAllBytes($path)
        $bytes[0] | Should -Be 0xFF
        $bytes[1] | Should -Be 0xFE
        $text = [System.Text.Encoding]::Unicode.GetString($bytes)
        $text | Should -Match ([regex]::Escape("$firstName,1"))
        $text | Should -Match ([regex]::Escape("$secondName,2"))
    }

    It 'appends CLR object rows using existing header casing insensitively' {
        $path = Join-Path $TestDrive 'append-clr-case.csv'
        Set-Content -LiteralPath $path -Value "name,value`nAlpha,1" -Encoding UTF8
        $type = 'PSWriteOffice.Tests.CsvAppendCaseRow' -as [type]
        if (-not $type) {
            $source = "namespace PSWriteOffice.Tests {`n    public sealed class CsvAppendCaseRow {`n        public string Name { get; set; }`n        public int Value { get; set; }`n    }`n}"

            Add-Type -TypeDefinition $source
            $type = 'PSWriteOffice.Tests.CsvAppendCaseRow' -as [type]
        }

        $row = [Activator]::CreateInstance($type)
        $row.Name = 'Beta'
        $row.Value = 2

        $row | Export-OfficeCsv -Path $path -Append

        Get-Content -LiteralPath $path | Should -Be @('name,value', 'Alpha,1', 'Beta,2')
    }

    It 'preserves DataRow adapter columns when converting CSV' {
        $table = [System.Data.DataTable]::new()
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Alpha', 1)

        $csvText = @($table.Rows | ConvertTo-OfficeCsv)

        $csvText | Should -Be @('Name,Value', 'Alpha,1')
    }

    It 'requires existing append columns unless Force is specified' {
        $path = Join-Path $TestDrive 'append-force.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        {
            [pscustomobject]@{ Name = 'Beta' } |
                Export-OfficeCsv -Path $path -Append -ErrorAction Stop
        } | Should -Throw

        [pscustomobject]@{ Name = 'Beta' } |
            Export-OfficeCsv -Path $path -Append -Force

        $data = Import-OfficeCsv -Path $path
        $data.Count | Should -Be 2
        $data[1].Name | Should -Be 'Beta'
        $data[1].Value | Should -Be ''
    }

    It 'projects forced scalar appends into an existing Value column' {
        $path = Join-Path $TestDrive 'append-force-scalar.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        'Beta' | Export-OfficeCsv -Path $path -Append -Force

        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1', ',Beta')
        $data = Import-OfficeCsv -Path $path
        $data[1].Name | Should -Be ''
        $data[1].Value | Should -Be 'Beta'
    }

    It 'projects forced no-header scalar appends into an existing Value column' {
        $path = Join-Path $TestDrive 'append-force-scalar-no-header.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        'Beta' | Export-OfficeCsv -Path $path -Append -NoHeader -Force

        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1', ',Beta')
        $data = Import-OfficeCsv -Path $path
        $data[1].Name | Should -Be ''
        $data[1].Value | Should -Be 'Beta'
    }

    It 'keeps forced scalar appends on the existing CSV schema without a Value column' {
        $path = Join-Path $TestDrive 'append-force-scalar-schema.csv'
        [pscustomobject]@{ Name = 'Alpha'; Other = 'One' } |
            Export-OfficeCsv -Path $path

        'Beta' | Export-OfficeCsv -Path $path -Append -Force

        Get-Content -LiteralPath $path | Should -Be @('Name,Other', 'Alpha,One', ',')
        $data = @(Import-OfficeCsv -Path $path)
        $data.Count | Should -Be 2
        $data[1].Name | Should -Be ''
        $data[1].Other | Should -Be ''
    }

    It 'keeps mixed scalar conversions on the first row schema without a Value column' {
        $csvText = @(
            [pscustomobject]@{ Name = 'Alpha'; Other = 'One' }
            'Beta'
        ) | ConvertTo-OfficeCsv

        $csvText | Should -Be @('Name,Other', 'Alpha,One', ',')
    }

    It 'does not touch an append target when first row validation fails' {
        $path = Join-Path $TestDrive 'append-validation-preserve.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -NoNewline -Encoding UTF8
        $before = [System.IO.File]::ReadAllBytes($path)

        {
            [pscustomobject]@{ Name = 'Beta' } |
                Export-OfficeCsv -Path $path -Append -ErrorAction Stop
        } | Should -Throw '*missing*Value*'

        $after = [System.IO.File]::ReadAllBytes($path)
        [Convert]::ToBase64String($after) | Should -Be ([Convert]::ToBase64String($before))
    }

    It 'validates every appended row against existing columns unless Force is specified' {
        $path = Join-Path $TestDrive 'append-validate-every-row.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        {
            @(
                [pscustomobject]@{ Name = 'Beta'; Value = 2 }
                [pscustomobject]@{ Name = 'Gamma' }
            ) | Export-OfficeCsv -Path $path -Append -ErrorAction Stop
        } | Should -Throw '*missing*Value*'
    }

    It 'appends CSV documents without writing duplicate headers' {
        $path = Join-Path $TestDrive 'append-document.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        $document = Get-OfficeCsv -Text "Name,Value`nBeta,2"
        Export-OfficeCsv -Document $document -Path $path -Append

        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1', 'Beta,2')
    }

    It 'appends multiple piped CSV documents to the same path' {
        $path = Join-Path $TestDrive 'append-piped-documents.csv'
        $documents = @(
            Get-OfficeCsv -Text "Name,Value`nAlpha,1"
            Get-OfficeCsv -Text "Name,Value`nBeta,2"
        )

        $documents | Export-OfficeCsv -Path $path -Append

        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1', 'Beta,2')
    }

    It 'serializes piped CSV documents as rows instead of object properties' {
        $path = Join-Path $TestDrive 'piped-document.csv'
        $document = Get-OfficeCsv -Text "Name,Value`nAlpha,1"

        $csvText = @($document | ConvertTo-OfficeCsv)
        $document | Export-OfficeCsv -Path $path

        $csvText | Should -Be @('Name,Value', 'Alpha,1')
        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1')
    }

    It 'treats top-level Guid values as scalar CSV values' {
        $guid = [guid]::Parse('00112233-4455-6677-8899-aabbccddeeff')

        $csvText = @($guid | ConvertTo-OfficeCsv -NoHeader)

        $csvText | Should -Be '00112233-4455-6677-8899-aabbccddeeff'
    }

    It 'preserves ETS members added to scalar CSV values' {
        $value = 'Alpha' | Add-Member -NotePropertyName Name -NotePropertyValue 'Decorated' -PassThru

        $csvText = @($value | ConvertTo-OfficeCsv)

        $csvText[0] | Should -BeLike '*Name*'
        $csvText[1] | Should -BeLike '*Decorated*'
    }

    It 'uses the selected culture list separator when UseCulture is specified' {
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL')

        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            ConvertTo-OfficeCsv -UseCulture -Culture $culture

        $csvText | Should -Contain 'Name;Value'
    }

    It 'uses the selected culture list separator when reading CSV data' {
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL')
        $path = Join-Path $TestDrive 'culture-read.csv'
        Set-Content -LiteralPath $path -Value "Name;Value`nAlpha;1" -Encoding UTF8

        $data = @(Import-OfficeCsv -Path $path -UseCulture -Culture $culture)
        $row = $data | Select-Object -First 1

        $data.Count | Should -Be 1
        $row.GetType().FullName | Should -Be 'System.Management.Automation.PSCustomObject'
        $row.Name | Should -Be 'Alpha'
        $row.Value | Should -Be '1'
    }

    It 'detects delimiters for CSV documents and row output' {
        $path = Join-Path $TestDrive 'detect-read.csv'
        Set-Content -LiteralPath $path -Value "Field1;Field2;Field3`n1,2,3,4;5,6,7,8;9,10,11,12" -Encoding UTF8

        $document = Get-OfficeCsv -Path $path -DetectDelimiter
        $data = Import-OfficeCsv -Path $path -DetectDelimiter

        $document.Delimiter | Should -Be ';'
        $document.Header | Should -Be @('Field1', 'Field2', 'Field3')
        $data[0].Field2 | Should -Be '5,6,7,8'
    }

    It 'detects delimiters after skipped preamble rows' {
        $path = Join-Path $TestDrive 'detect-after-preamble.csv'
        Set-Content -LiteralPath $path -Value "generated,by,vendor,with,commas`nName;Value`nAlpha;1" -Encoding UTF8

        $document = Get-OfficeCsv -Path $path -DetectDelimiter -SkipRows 1
        $data = Import-OfficeCsv -Path $path -DetectDelimiter -SkipRows 1

        $document.Delimiter | Should -Be ';'
        $document.Header | Should -Be @('Name', 'Value')
        $data[0].Value | Should -Be '1'
    }

    It 'uses delimiter candidates when detecting from text' {
        $document = Get-OfficeCsv -Text "Name|Value`nAlpha|1" -DetectDelimiter -DelimiterCandidates ';', '|'

        $document.Delimiter | Should -Be '|'
        $row = @($document.AsEnumerable())[0]
        $row['Value'] | Should -Be '1'
    }

    It 'converts CSV text directly into row objects' {
        $data = @(ConvertFrom-OfficeCsv -Text "Name|Value`nAlpha|1" -DetectDelimiter -DelimiterCandidates ';', '|')

        $data.Count | Should -Be 1
        $data[0].Name | Should -Be 'Alpha'
        $data[0].Value | Should -Be '1'
    }

    It 'parses piped CSV text as one stream' {
        $data = "Name,Value", "Alpha,1", "Beta,2" | ConvertFrom-OfficeCsv

        $data.Count | Should -Be 2
        $data[0].Name | Should -Be 'Alpha'
        $data[0].Value | Should -Be '1'
        $data[1].Name | Should -Be 'Beta'
        $data[1].Value | Should -Be '2'
    }

    It 'imports piped file paths as paths rather than CSV text' {
        $path = Join-Path $TestDrive 'piped-path.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -Encoding UTF8

        $data = @($path | Import-OfficeCsv)

        $data.Count | Should -Be 1
        $data[0].Name | Should -Be 'Alpha'
    }

    It 'loads CSV documents from literal paths' {
        $path = Join-Path $TestDrive 'literal[1].csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -Encoding UTF8

        $document = Get-OfficeCsv -LiteralPath $path

        $document.Header | Should -Be @('Name', 'Value')
        $document.AsEnumerable().Count | Should -Be 1
    }

    It 'loads CSV documents using the legacy InputPath alias' {
        $path = Join-Path $TestDrive 'input-path-alias.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -Encoding UTF8

        $document = Get-OfficeCsv -InputPath $path

        $document.Header | Should -Be @('Name', 'Value')
        $document.AsEnumerable().Count | Should -Be 1
    }

    It 'expands Path wildcards when importing CSV rows' {
        $folder = Join-Path $TestDrive 'wildcard-import'
        New-Item -Path $folder -ItemType Directory | Out-Null
        Set-Content -LiteralPath (Join-Path $folder 'a.csv') -Value "Name,Value`nAlpha,1" -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $folder 'b.csv') -Value "Name,Value`nBeta,2" -Encoding UTF8

        $data = Import-OfficeCsv -Path (Join-Path $folder '*.csv') | Sort-Object Name

        $data.Count | Should -Be 2
        $data[0].Name | Should -Be 'Alpha'
        $data[1].Name | Should -Be 'Beta'
    }

    It 'loads multiple CSV documents from Path values' {
        $folder = Join-Path $TestDrive 'multi-document'
        New-Item -Path $folder -ItemType Directory | Out-Null
        $paths = @(
            Join-Path $folder 'first.csv'
            Join-Path $folder 'second.csv'
        )
        Set-Content -LiteralPath $paths[0] -Value "Name,Value`nAlpha,1" -Encoding UTF8
        Set-Content -LiteralPath $paths[1] -Value "Name,Value`nBeta,2" -Encoding UTF8

        $documents = @(Get-OfficeCsv -Path $paths)

        $documents.Count | Should -Be 2
        $documents[0].Header | Should -Be @('Name', 'Value')
        @($documents[1].AsEnumerable())[0]['Name'] | Should -Be 'Beta'
    }

    It 'preserves unquoted whitespace by default and trims when requested' {
        $path = Join-Path $TestDrive 'whitespace.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,  spaced  " -Encoding UTF8

        $default = Import-OfficeCsv -Path $path
        $trimmed = Import-OfficeCsv -Path $path -TrimWhitespace:$true

        $default[0].Value | Should -Be '  spaced  '
        $trimmed[0].Value | Should -Be 'spaced'
    }

    It 'uses explicit headers and treats the first row as data' {
        $path = Join-Path $TestDrive 'explicit-header.csv'
        Set-Content -LiteralPath $path -Value "Alpha,1`nBeta,2" -Encoding UTF8

        $data = Import-OfficeCsv -Path $path -Header Name, Value

        $data.Count | Should -Be 2
        $data[0].Name | Should -Be 'Alpha'
        $data[1].Value | Should -Be '2'
    }

    It 'rejects duplicate object headers before using the optimized output path' {
        $path = Join-Path $TestDrive 'duplicate-header.csv'
        Set-Content -LiteralPath $path -Value "Name,Name`nAlpha,1" -Encoding UTF8

        { Import-OfficeCsv -Path $path -ErrorAction Stop } | Should -Throw
    }

    It 'rejects duplicate hashtable headers instead of overwriting values' {
        $path = Join-Path $TestDrive 'duplicate-hashtable-header.csv'
        Set-Content -LiteralPath $path -Value "Name,Name`nAlpha,1" -Encoding UTF8

        { Import-OfficeCsv -Path $path -AsHashtable -ErrorAction Stop } | Should -Throw
    }

    It 'supports NoHeader when reading CSV data and documents' {
        $path = Join-Path $TestDrive 'no-header-read.csv'
        Set-Content -LiteralPath $path -Value "Alpha,1`nBeta,2" -Encoding UTF8

        $data = Import-OfficeCsv -Path $path -NoHeader
        $document = Get-OfficeCsv -Path $path -NoHeader

        $data.Count | Should -Be 2
        $data[0].Column1 | Should -Be 'Alpha'
        $data[0].Column2 | Should -Be '1'
        $document.Header | Should -Be @('Column1', 'Column2')
        @($document.AsEnumerable()).Count | Should -Be 2
    }

    It 'skips initial records before CSV header discovery' {
        $path = Join-Path $TestDrive 'skip-rows.csv'
        Set-Content -LiteralPath $path -Value "generated by vendor`nexported today`nName,Value`nAlpha,1" -Encoding UTF8

        $data = @(Import-OfficeCsv -Path $path -SkipRows 2)
        $document = Get-OfficeCsv -Path $path -SkipRows 2
        $fromText = @(ConvertFrom-OfficeCsv -Text "metadata`nName,Value`nBeta,2" -SkipRows 1)

        $data.Count | Should -Be 1
        $data[0].Name | Should -Be 'Alpha'
        $document.Header | Should -Be @('Name', 'Value')
        $fromText[0].Name | Should -Be 'Beta'
    }

    It 'generates missing header names and tolerates uneven rows by default' {
        $path = Join-Path $TestDrive 'uneven.csv'
        Set-Content -LiteralPath $path -Value "Name,,Value`nAlpha,Ignored`nBeta,Ignored,2,Extra" -Encoding UTF8

        $data = @(Import-OfficeCsv -Path $path)
        $rows = @($data | Select-Object -First 2)

        $data.Count | Should -Be 2
        $rows[0].GetType().FullName | Should -Be 'System.Management.Automation.PSCustomObject'
        $rows[0].H1 | Should -Be 'Ignored'
        $rows[0].Value | Should -Be ''
        $rows[1].Value | Should -Be '2'
    }

    It 'keeps trailing missing header properties consistent across import modes' {
        $path = Join-Path $TestDrive 'padded.csv'
        Set-Content -LiteralPath $path -Value "Name,Value,Other`nAlpha,1" -Encoding UTF8

        $streamed = @(Import-OfficeCsv -Path $path)
        $inMemory = @(Import-OfficeCsv -Path $path -Mode InMemory)
        $fromDocument = @(Get-OfficeCsv -Path $path | Import-OfficeCsv)

        foreach ($row in @($streamed[0], $inMemory[0], $fromDocument[0])) {
            $row.PSObject.Properties.Name | Should -Contain 'Other'
            $row.Other | Should -BeNullOrEmpty
        }
    }

    It 'can enforce strict row width validation' {
        $path = Join-Path $TestDrive 'strict-uneven.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha" -Encoding UTF8

        { Import-OfficeCsv -Path $path -ColumnCountMismatchPolicy Strict } | Should -Throw
    }

    It 'recognizes W3C fields headers when loading CSV documents' {
        $document = Get-OfficeCsv -Text "#Version: 1.0`n#Fields: date time cs-uri`n2026-06-24 12:00 /index" -Delimiter ' '

        $document.Header | Should -Be @('date', 'time', 'cs-uri')
        $row = @($document.AsEnumerable())[0]
        $row['cs-uri'] | Should -Be '/index'
    }

    It 'can treat a leading comment row as the header when requested' {
        $document = Get-OfficeCsv -Text "#Name,Value`nAlpha,1" -SkipCommentRowsBeforeHeader:$false

        $document.Header | Should -Be @('#Name', 'Value')
        $row = @($document.AsEnumerable())[0]
        $row['#Name'] | Should -Be 'Alpha'
    }

    It 'does not treat quoted comment-character headers as comments' {
        $path = Join-Path $TestDrive 'quoted-comment-header.csv'
        Set-Content -LiteralPath $path -Value '"#Tag",Name', '10,Alpha' -Encoding UTF8

        $data = @(Import-OfficeCsv -Path $path)
        $document = Get-OfficeCsv -Path $path

        $document.Header | Should -Be @('#Tag', 'Name')
        $data.Count | Should -Be 1
        $data[0].'#Tag' | Should -Be '10'
        $data[0].Name | Should -Be 'Alpha'
    }

    It 'can skip comment rows throughout the file' {
        $path = Join-Path $TestDrive 'comments.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`n# ignored`nBeta,2" -Encoding UTF8

        $data = Import-OfficeCsv -Path $path -SkipCommentRows

        $data.Count | Should -Be 2
        $data[0].Name | Should -Be 'Alpha'
        $data[1].Name | Should -Be 'Beta'
    }

    It 'can skip custom comment rows throughout the file' {
        $path = Join-Path $TestDrive 'custom-comments.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`n; ignored`nBeta,2" -Encoding UTF8

        $data = Import-OfficeCsv -Path $path -SkipCommentRows -CommentCharacter ';'

        $data.Count | Should -Be 2
        $data[1].Name | Should -Be 'Beta'
    }

    It 'lets parameter binding reject Delimiter and UseCulture together' {
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL')

        {
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
                ConvertTo-OfficeCsv -Delimiter ',' -UseCulture -Culture $culture
        } | Should -Throw
    }

    It 'lets parameter binding reject Delimiter and DetectDelimiter together' {
        {
            Get-OfficeCsv -Text "Name;Value`nAlpha;1" -Delimiter ';' -DetectDelimiter
        } | Should -Throw
    }

    It 'rejects Header and NoHeader together on CSV read surfaces' {
        $path = Join-Path $TestDrive 'header-noheader.csv'
        Set-Content -LiteralPath $path -Value "Alpha,1" -Encoding UTF8

        { Import-OfficeCsv -Path $path -Header Name, Value -NoHeader -ErrorAction Stop } | Should -Throw '*Header*NoHeader*'
        { Get-OfficeCsv -Path $path -Header Name, Value -NoHeader -ErrorAction Stop } | Should -Throw '*Header*NoHeader*'
        { ConvertFrom-OfficeCsv -Text "Alpha,1" -Header Name, Value -NoHeader -ErrorAction Stop } | Should -Throw '*Header*NoHeader*'
    }

    It 'keeps file-only encoding off text parameter sets' {
        $textSets = (Get-Command Get-OfficeCsv).ParameterSets |
            Where-Object Name -like 'Text*'

        foreach ($set in $textSets) {
            $set.Parameters.Name | Should -Not -Contain 'Encoding'
        }

        (Get-Command ConvertTo-OfficeCsv).Parameters.Keys | Should -Not -Contain 'Encoding'
    }

    It 'streams ConvertTo-OfficeCsv output as CSV records that ConvertFrom-OfficeCsv can read' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $csvLines = @($rows | ConvertTo-OfficeCsv)
        $roundTrip = $csvLines | ConvertFrom-OfficeCsv

        $csvLines | Should -Be @('Name,Value', 'Alpha,1', 'Beta,2')
        $roundTrip.Count | Should -Be 2
        $roundTrip[1].Name | Should -Be 'Beta'
        $roundTrip[1].Value | Should -Be '2'
    }

    It 'keeps quoted embedded newlines inside one ConvertTo-OfficeCsv record object' {
        $csvLines = @([pscustomobject]@{ Name = 'Alpha'; Note = "one`ntwo" } | ConvertTo-OfficeCsv)

        $csvLines.Count | Should -Be 2
        $csvLines[0] | Should -Be 'Name,Note'
        $csvLines[1] | Should -Be "Alpha,`"one`ntwo`""
        @($csvLines | ConvertFrom-OfficeCsv)[0].Note | Should -Be "one`ntwo"
    }

    It 'keeps separate records when unquoted values contain quote characters' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Note = 'a"b' }
            [pscustomobject]@{ Name = 'Beta'; Note = 'plain' }
        )

        $csvLines = @($rows | ConvertTo-OfficeCsv -UseQuotes Never)

        $csvLines.Count | Should -Be 3
        $csvLines | Should -Be @('Name,Note', 'Alpha,a"b', 'Beta,plain')
    }

    It 'keeps separate records when unquoted values start with quote characters' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Note = '"starts' }
            [pscustomobject]@{ Name = 'Beta'; Note = 'plain' }
        )

        $csvLines = @($rows | ConvertTo-OfficeCsv -UseQuotes Never)

        $csvLines.Count | Should -Be 3
        $csvLines | Should -Be @('Name,Note', 'Alpha,"starts', 'Beta,plain')
    }

    It 'normalizes CLR projection values and skips failing CLR getters' {
        if (-not ('PSWriteOffice.Tests.CsvClrProjectionRow' -as [type])) {
            Add-Type -TypeDefinition @'
namespace PSWriteOffice.Tests {
    using System;

    public sealed class CsvClrProjectionRow {
        public string Name { get { return "Alpha"; } }
        public string Broken { get { throw new InvalidOperationException("boom"); } }
        public string[] Tags { get { return new[] { "one", "two" }; } }
    }
}
'@
        }

        $row = [PSWriteOffice.Tests.CsvClrProjectionRow]::new()
        $csvLines = @($row | ConvertTo-OfficeCsv)

        $csvLines | Should -Be @('Name,Tags', 'Alpha,"one, two"')
    }

    It 'lets QuoteFields compose with UseQuotes' {
        $path = Join-Path $TestDrive 'quoted-fields.csv'

        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = 1; Note = 'plain' } |
            ConvertTo-OfficeCsv -UseQuotes AsNeeded -QuoteFields Name

        [pscustomobject]@{ Name = 'Alpha'; Value = 1; Note = 'plain' } |
            Export-OfficeCsv -Path $path -UseQuotes AsNeeded -QuoteFields Name

        $csvText | Should -Contain '"Name",Value,Note'
        $csvText | Should -Contain '"Alpha",1,plain'
        (Get-Content -LiteralPath $path -Raw) | Should -Match '"Alpha",1,plain'
    }

    It 'escapes formula-like values when requested' {
        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = '=1+1' } |
            ConvertTo-OfficeCsv -FormulaInjectionPolicy Escape

        ($csvText -join "`n") | Should -Match "'=1\+1"
    }

    It 'uses AsNeeded quoting by default and supports PowerShell-style quote policies' {
        $row = [pscustomobject]@{ Name = 'Alpha'; Value = 'A,B'; Note = 'plain' }

        $default = $row | ConvertTo-OfficeCsv
        $always = $row | ConvertTo-OfficeCsv -UseQuotes Always
        $never = $row | ConvertTo-OfficeCsv -UseQuotes Never
        $quoteFields = $row | ConvertTo-OfficeCsv -QuoteFields Name, Note

        $default | Should -Contain 'Name,Value,Note'
        $default | Should -Contain 'Alpha,"A,B",plain'
        $always | Should -Contain '"Name","Value","Note"'
        $always | Should -Contain '"Alpha","A,B","plain"'
        $never | Should -Contain 'Alpha,A,B,plain'
        $quoteFields | Should -Contain '"Name",Value,"Note"'
        $quoteFields | Should -Contain '"Alpha","A,B","plain"'
    }

    It 'quotes empty values when the quote policy is Always' {
        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = $null } |
            ConvertTo-OfficeCsv -UseQuotes Always

        $csvText | Should -Contain '"Alpha",""'
    }

    It 'supports NoHeader when converting and exporting CSV' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )
        $path = Join-Path $TestDrive 'no-header-export.csv'

        $csvText = $rows | ConvertTo-OfficeCsv -NoHeader
        $rows | Export-OfficeCsv -Path $path -NoHeader

        ($csvText -join "`n") | Should -Not -Match 'Name'
        $csvText | Should -Contain 'Alpha,1'
        (Get-Content -LiteralPath $path -Raw) | Should -Not -Match 'Name'
    }
}
