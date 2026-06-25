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

        (Get-Command ConvertTo-OfficeCsv).Parameters.Keys | Should -Not -Contain 'IncludeHeader'
        (Get-Command Export-OfficeCsv).Parameters.Keys | Should -Not -Contain 'IncludeHeader'
        (Get-Command Get-OfficeCsv).Parameters.Keys | Should -Not -Contain 'HasHeaderRow'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Not -Contain 'HasHeaderRow'
    }

    It 'exposes CSV row import with idiomatic command names' {
        (Get-Command Import-OfficeCsv).CommandType | Should -Be 'Cmdlet'
        (Get-Command Get-OfficeCsvData).ResolvedCommandName | Should -Be 'Import-OfficeCsv'
        (Get-Command ConvertFrom-OfficeCsv).CommandType | Should -Be 'Cmdlet'
    }

    It 'converts objects to CSV and reads them back' {
        $rows = @(
            [pscustomobject]@{ Region = 'NA'; Revenue = 100 }
            [pscustomobject]@{ Region = 'EMEA'; Revenue = 200 }
        )

        $csvText = $rows | ConvertTo-OfficeCsv
        $csvText | Should -Match 'Region'
        $csvText | Should -Match 'Revenue'

        $path = Join-Path $TestDrive 'data.csv'
        $rows | Export-OfficeCsv -Path $path | Out-Null

        Test-Path $path | Should -BeTrue

        $data = Import-OfficeCsv -Path $path
        $data.Count | Should -Be 2
        $data[0].Region | Should -Be 'NA'
    }

    It 'writes to files using the Path alias' {
        $path = Join-Path $TestDrive 'path-alias.csv'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        Test-Path $path | Should -BeTrue
        Get-Content -Path $path -Raw | Should -Match 'Alpha'
    }

    It 'uses the selected culture list separator when UseCulture is specified' {
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL')

        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            ConvertTo-OfficeCsv -UseCulture -Culture $culture

        $csvText | Should -Match 'Name;Value'
    }

    It 'uses the selected culture list separator when reading CSV data' {
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL')
        $path = Join-Path $TestDrive 'culture-read.csv'
        Set-Content -LiteralPath $path -Value "Name;Value`nAlpha;1" -Encoding UTF8

        $data = Import-OfficeCsv -Path $path -UseCulture -Culture $culture

        $data.Count | Should -Be 1
        $data[0].GetType().FullName | Should -Be 'System.Management.Automation.PSCustomObject'
        $data[0].Name | Should -Be 'Alpha'
        $data[0].Value | Should -Be '1'
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

    It 'uses delimiter candidates when detecting from text' {
        $document = Get-OfficeCsv -Text "Name|Value`nAlpha|1" -DetectDelimiter -DelimiterCandidates ';', '|'

        $document.Delimiter | Should -Be '|'
        $row = @($document.AsEnumerable())[0]
        $row['Value'] | Should -Be '1'
    }

    It 'converts CSV text directly into row objects' {
        $data = ConvertFrom-OfficeCsv -Text "Name|Value`nAlpha|1" -DetectDelimiter -DelimiterCandidates ';', '|'

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

        $data = $path | Import-OfficeCsv

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

    It 'generates missing header names and tolerates uneven rows by default' {
        $path = Join-Path $TestDrive 'uneven.csv'
        Set-Content -LiteralPath $path -Value "Name,,Value`nAlpha,Ignored`nBeta,Ignored,2,Extra" -Encoding UTF8

        $data = Import-OfficeCsv -Path $path

        $data.Count | Should -Be 2
        $data[0].GetType().FullName | Should -Be 'System.Management.Automation.PSCustomObject'
        $data[0].H1 | Should -Be 'Ignored'
        $data[0].Value | Should -Be ''
        $data[1].Value | Should -Be '2'
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

    It 'keeps file-only encoding off text parameter sets' {
        $textSets = (Get-Command Get-OfficeCsv).ParameterSets |
            Where-Object Name -like 'Text*'

        foreach ($set in $textSets) {
            $set.Parameters.Name | Should -Not -Contain 'Encoding'
        }
    }

    It 'lets QuoteFields compose with UseQuotes' {
        $path = Join-Path $TestDrive 'quoted-fields.csv'

        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = 1; Note = 'plain' } |
            ConvertTo-OfficeCsv -UseQuotes AsNeeded -QuoteFields Name

        [pscustomobject]@{ Name = 'Alpha'; Value = 1; Note = 'plain' } |
            Export-OfficeCsv -Path $path -UseQuotes AsNeeded -QuoteFields Name

        $csvText | Should -Match '"Name",Value,Note'
        $csvText | Should -Match '"Alpha",1,plain'
        (Get-Content -LiteralPath $path -Raw) | Should -Match '"Alpha",1,plain'
    }

    It 'escapes formula-like values when requested' {
        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = '=1+1' } |
            ConvertTo-OfficeCsv -FormulaInjectionPolicy Escape

        $csvText | Should -Match "'=1\+1"
    }

    It 'uses AsNeeded quoting by default and supports PowerShell-style quote policies' {
        $row = [pscustomobject]@{ Name = 'Alpha'; Value = 'A,B'; Note = 'plain' }

        $default = $row | ConvertTo-OfficeCsv
        $always = $row | ConvertTo-OfficeCsv -UseQuotes Always
        $never = $row | ConvertTo-OfficeCsv -UseQuotes Never
        $quoteFields = $row | ConvertTo-OfficeCsv -QuoteFields Name, Note

        $default | Should -Match 'Name,Value,Note'
        $default | Should -Match 'Alpha,"A,B",plain'
        $always | Should -Match '"Name","Value","Note"'
        $always | Should -Match '"Alpha","A,B","plain"'
        $never | Should -Match 'Alpha,A,B,plain'
        $quoteFields | Should -Match '"Name",Value,"Note"'
        $quoteFields | Should -Match '"Alpha","A,B","plain"'
    }

    It 'quotes empty values when the quote policy is Always' {
        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = $null } |
            ConvertTo-OfficeCsv -UseQuotes Always

        $csvText | Should -Match '"Alpha",""'
    }

    It 'supports NoHeader when converting and exporting CSV' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )
        $path = Join-Path $TestDrive 'no-header-export.csv'

        $csvText = $rows | ConvertTo-OfficeCsv -NoHeader
        $rows | Export-OfficeCsv -Path $path -NoHeader

        $csvText | Should -Not -Match 'Name'
        $csvText | Should -Match 'Alpha,1'
        (Get-Content -LiteralPath $path -Raw) | Should -Not -Match 'Name'
    }
}
