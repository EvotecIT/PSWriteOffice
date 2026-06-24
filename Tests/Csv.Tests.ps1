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

        $data = Get-OfficeCsvData -Path $path
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

        $csvText | Should -Match '"Name";"Value"'
    }

    It 'uses the selected culture list separator when reading CSV data' {
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL')
        $path = Join-Path $TestDrive 'culture-read.csv'
        Set-Content -LiteralPath $path -Value "Name;Value`nAlpha;1" -Encoding UTF8

        $data = Get-OfficeCsvData -Path $path -UseCulture -Culture $culture

        $data.Count | Should -Be 1
        $data[0].Name | Should -Be 'Alpha'
        $data[0].Value | Should -Be '1'
    }

    It 'detects delimiters for CSV documents and row output' {
        $path = Join-Path $TestDrive 'detect-read.csv'
        Set-Content -LiteralPath $path -Value "Field1;Field2;Field3`n1,2,3,4;5,6,7,8;9,10,11,12" -Encoding UTF8

        $document = Get-OfficeCsv -Path $path -DetectDelimiter
        $data = Get-OfficeCsvData -Path $path -DetectDelimiter

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

    It 'loads CSV documents from literal paths' {
        $path = Join-Path $TestDrive 'literal[1].csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -Encoding UTF8

        $document = Get-OfficeCsv -LiteralPath $path

        $document.Header | Should -Be @('Name', 'Value')
        $document.AsEnumerable().Count | Should -Be 1
    }

    It 'preserves unquoted whitespace by default and trims when requested' {
        $path = Join-Path $TestDrive 'whitespace.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,  spaced  " -Encoding UTF8

        $default = Get-OfficeCsvData -Path $path
        $trimmed = Get-OfficeCsvData -Path $path -TrimWhitespace:$true

        $default[0].Value | Should -Be '  spaced  '
        $trimmed[0].Value | Should -Be 'spaced'
    }

    It 'uses explicit headers and treats the first row as data' {
        $path = Join-Path $TestDrive 'explicit-header.csv'
        Set-Content -LiteralPath $path -Value "Alpha,1`nBeta,2" -Encoding UTF8

        $data = Get-OfficeCsvData -Path $path -Header Name, Value

        $data.Count | Should -Be 2
        $data[0].Name | Should -Be 'Alpha'
        $data[1].Value | Should -Be '2'
    }

    It 'generates missing header names and tolerates uneven rows by default' {
        $path = Join-Path $TestDrive 'uneven.csv'
        Set-Content -LiteralPath $path -Value "Name,,Value`nAlpha,Ignored`nBeta,Ignored,2,Extra" -Encoding UTF8

        $data = Get-OfficeCsvData -Path $path

        $data.Count | Should -Be 2
        $data[0].H1 | Should -Be 'Ignored'
        $data[0].Value | Should -Be ''
        $data[1].Value | Should -Be '2'
    }

    It 'can enforce strict row width validation' {
        $path = Join-Path $TestDrive 'strict-uneven.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha" -Encoding UTF8

        { Get-OfficeCsvData -Path $path -ColumnCountMismatchPolicy Strict } | Should -Throw
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

        $data = Get-OfficeCsvData -Path $path -SkipCommentRows

        $data.Count | Should -Be 2
        $data[0].Name | Should -Be 'Alpha'
        $data[1].Name | Should -Be 'Beta'
    }

    It 'can skip custom comment rows throughout the file' {
        $path = Join-Path $TestDrive 'custom-comments.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`n; ignored`nBeta,2" -Encoding UTF8

        $data = Get-OfficeCsvData -Path $path -SkipCommentRows -CommentCharacter ';'

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

    It 'lets parameter binding reject UseQuotes and QuoteFields together' {
        {
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
                ConvertTo-OfficeCsv -UseQuotes AsNeeded -QuoteFields Name
        } | Should -Throw
    }

    It 'escapes formula-like values when requested' {
        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = '=1+1' } |
            ConvertTo-OfficeCsv -FormulaInjectionPolicy Escape

        $csvText | Should -Match "'=1\+1"
    }

    It 'supports PowerShell-style quote policies and selected quote fields' {
        $row = [pscustomobject]@{ Name = 'Alpha'; Value = 'A,B'; Note = 'plain' }

        $asNeeded = $row | ConvertTo-OfficeCsv -UseQuotes AsNeeded
        $always = $row | ConvertTo-OfficeCsv
        $never = $row | ConvertTo-OfficeCsv -UseQuotes Never
        $quoteFields = $row | ConvertTo-OfficeCsv -QuoteFields Name, Note

        $asNeeded | Should -Match 'Alpha,"A,B",plain'
        $always | Should -Match '"Name","Value","Note"'
        $always | Should -Match '"Alpha","A,B","plain"'
        $never | Should -Match 'Alpha,A,B,plain'
        $quoteFields | Should -Match '"Name",Value,"Note"'
        $quoteFields | Should -Match '"Alpha","A,B","plain"'
    }

    It 'quotes empty values when the quote policy is Always' {
        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = $null } |
            ConvertTo-OfficeCsv

        $csvText | Should -Match '"Alpha",""'
    }
}
