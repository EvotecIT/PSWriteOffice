BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Reader cmdlets' {
    It 'exposes built-in and modular Reader capabilities' {
        $capabilities = Get-OfficeDocumentCapability

        $capabilities.Id | Should -Contain 'officeimo.reader.word'
        $capabilities.Id | Should -Contain 'officeimo.reader.excel'
        $capabilities.Id | Should -Contain 'officeimo.reader.powerpoint'
        $capabilities.Id | Should -Contain 'officeimo.reader.pdf'
        $capabilities.Id | Should -Contain 'officeimo.reader.html'
        $capabilities.Id | Should -Contain 'officeimo.reader.csv'
        $capabilities.Id | Should -Contain 'officeimo.reader.json'
        $capabilities.Id | Should -Contain 'officeimo.reader.xml'
        $capabilities.Id | Should -Contain 'officeimo.reader.yaml'
        $capabilities.Id | Should -Contain 'officeimo.reader.zip'
        $capabilities.Id | Should -Contain 'officeimo.reader.epub'
        $capabilities.Id | Should -Contain 'officeimo.reader.visio'
    }

    It 'exports Reader table, visual, and ingestion cmdlets' {
        Get-Command -Name Get-OfficeDocumentTable -ErrorAction Stop | Should -Not -BeNullOrEmpty
        Get-Command -Name Get-OfficeDocumentVisual -ErrorAction Stop | Should -Not -BeNullOrEmpty
        Get-Command -Name Get-OfficeDocumentIngest -ErrorAction Stop | Should -Not -BeNullOrEmpty
        Get-Command -Name Read-OfficeDocumentTable -ErrorAction Stop | Should -Not -BeNullOrEmpty
        Get-Command -Name Read-OfficeDocumentVisual -ErrorAction Stop | Should -Not -BeNullOrEmpty
    }

    It 'does not replace caller-registered PDF readers' {
        $handlerId = 'pswriteoffice.test.pdf'
        $documentReaderType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader' -TypeName 'OfficeIMO.Reader.DocumentReader' -CommandName 'Get-OfficeDocumentCapability'
        $registrationType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader' -TypeName 'OfficeIMO.Reader.ReaderHandlerRegistration' -CommandName 'Get-OfficeDocumentCapability'
        $inputKindType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader' -TypeName 'OfficeIMO.Reader.ReaderInputKind' -CommandName 'Get-OfficeDocumentCapability'
        $chunkType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader' -TypeName 'OfficeIMO.Reader.ReaderChunk' -CommandName 'Get-OfficeDocumentCapability'
        $pdfRegistrationType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader.Pdf' -TypeName 'OfficeIMO.Reader.Pdf.DocumentReaderPdfRegistrationExtensions' -CommandName 'Get-OfficeDocumentCapability'

        $unregisterHandler = $documentReaderType.GetMethod('UnregisterHandler', [type[]] @([string]))
        $registerHandler = $documentReaderType.GetMethod('RegisterHandler', [type[]] @($registrationType, [bool]))
        $unregisterPdfHandler = $pdfRegistrationType.GetMethod('UnregisterPdfHandler', [System.Reflection.BindingFlags]'Public, Static')
        $unregisterHandler.Invoke($null, @($handlerId)) | Out-Null

        $registration = [Activator]::CreateInstance($registrationType)
        $registration.Id = $handlerId
        $registration.DisplayName = 'Test PDF Reader'
        $registration.Kind = [Enum]::Parse($inputKindType, 'Pdf')
        $registration.Extensions = [string[]]@('.pdf')
        $readPathType = $registrationType.GetProperty('ReadPath').PropertyType
        $registration.ReadPath = [System.Management.Automation.LanguagePrimitives]::ConvertTo({
            param($Path, $Options, $CancellationToken)

            [Array]::CreateInstance($chunkType, 0)
        }.GetNewClosure(), $readPathType)

        try {
            $registerHandler.Invoke($null, @($registration, $true)) | Out-Null

            $capabilities = @(Get-OfficeDocumentCapability -ExcludeBuiltIn)
            ($capabilities | Where-Object Id -EQ $handlerId).Count | Should -Be 1
            ($capabilities | Where-Object Id -EQ 'officeimo.reader.pdf').Count | Should -Be 0
        } finally {
            $unregisterHandler.Invoke($null, @($handlerId)) | Out-Null
            $unregisterPdfHandler.Invoke($null, @()) | Out-Null
        }
    }

    It 'keeps non-conflicting adapter extensions when a caller owns one extension' {
        $handlerId = 'pswriteoffice.test.html'
        $documentReaderType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader' -TypeName 'OfficeIMO.Reader.DocumentReader' -CommandName 'Get-OfficeDocumentCapability'
        $registrationType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader' -TypeName 'OfficeIMO.Reader.ReaderHandlerRegistration' -CommandName 'Get-OfficeDocumentCapability'
        $inputKindType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader' -TypeName 'OfficeIMO.Reader.ReaderInputKind' -CommandName 'Get-OfficeDocumentCapability'
        $chunkType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader' -TypeName 'OfficeIMO.Reader.ReaderChunk' -CommandName 'Get-OfficeDocumentCapability'
        $htmlRegistrationType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader.Html' -TypeName 'OfficeIMO.Reader.Html.DocumentReaderHtmlRegistrationExtensions' -CommandName 'Get-OfficeDocumentCapability'

        $unregisterHandler = $documentReaderType.GetMethod('UnregisterHandler', [type[]] @([string]))
        $registerHandler = $documentReaderType.GetMethod('RegisterHandler', [type[]] @($registrationType, [bool]))
        $unregisterHtmlHandler = $htmlRegistrationType.GetMethod('UnregisterHtmlHandler', [System.Reflection.BindingFlags]'Public, Static')
        $unregisterHandler.Invoke($null, @($handlerId)) | Out-Null
        $unregisterHtmlHandler.Invoke($null, @()) | Out-Null

        $registration = [Activator]::CreateInstance($registrationType)
        $registration.Id = $handlerId
        $registration.DisplayName = 'Test HTML Reader'
        $registration.Kind = [Enum]::Parse($inputKindType, 'Html')
        $registration.Extensions = [string[]]@('.html')
        $readPathType = $registrationType.GetProperty('ReadPath').PropertyType
        $registration.ReadPath = [System.Management.Automation.LanguagePrimitives]::ConvertTo({
            param($Path, $Options, $CancellationToken)

            [Array]::CreateInstance($chunkType, 0)
        }.GetNewClosure(), $readPathType)

        try {
            $registerHandler.Invoke($null, @($registration, $true)) | Out-Null

            $capabilities = @(Get-OfficeDocumentCapability -ExcludeBuiltIn)
            $customCapability = $capabilities | Where-Object Id -EQ $handlerId
            $htmlCapability = $capabilities | Where-Object Id -EQ 'officeimo.reader.html'

            $customCapability | Should -HaveCount 1
            $htmlCapability | Should -HaveCount 1
            $customCapability.Extensions | Should -Contain '.html'
            $htmlCapability.Extensions | Should -Not -Contain '.html'
            $htmlCapability.Extensions | Should -Contain '.htm'
            $htmlCapability.Extensions | Should -Contain '.xhtml'
        } finally {
            $unregisterHandler.Invoke($null, @($handlerId)) | Out-Null
            $unregisterHtmlHandler.Invoke($null, @()) | Out-Null
        }
    }

    It 'reads Markdown files as chunks and a document envelope' {
        $path = Join-Path $TestDrive 'source.md'
        Set-Content -Path $path -Value "# Reader smoke`n`nOfficeIMO Reader keeps this text." -Encoding UTF8

        $chunks = @(Get-OfficeDocumentChunk -Path $path)
        $chunks.Count | Should -BeGreaterThan 0
        ($chunks.Text -join "`n") | Should -Match 'OfficeIMO Reader keeps this text'

        $document = Get-OfficeDocument -Path $path
        $document.Chunks.Count | Should -BeGreaterThan 0
        $document.Markdown | Should -Match 'Reader smoke'

        $json = Get-OfficeDocument -Path $path -AsJson
        $json | Should -Match 'officeimo.document.read-result'
        $json | Should -Match 'OfficeIMO Reader keeps this text'
    }

    It 'reads HTML, JSON, and YAML files through modular Reader adapters' {
        $htmlPath = Join-Path $TestDrive 'source.html'
        Set-Content -Path $htmlPath -Value '<!doctype html><html><body><h1>Reader HTML</h1><p>OfficeIMO adapter text.</p></body></html>' -Encoding UTF8

        $htmlChunks = @(Get-OfficeDocumentChunk -Path $htmlPath)
        $htmlChunks.Count | Should -BeGreaterThan 0
        $htmlChunks[0].Kind.ToString() | Should -Be 'Html'
        ($htmlChunks.Text -join "`n") | Should -Match 'Reader HTML'
        ($htmlChunks.Text -join "`n") | Should -Match 'OfficeIMO adapter text'

        $jsonPath = Join-Path $TestDrive 'source.json'
        Set-Content -Path $jsonPath -Value '{"name":"Ada","score":42}' -Encoding UTF8

        $jsonChunks = @(Get-OfficeDocumentChunk -Path $jsonPath)
        $jsonChunks.Count | Should -BeGreaterThan 0
        $jsonChunks[0].Kind.ToString() | Should -Be 'Json'
        ($jsonChunks.Text -join "`n") | Should -Match '\$\.name'
        ($jsonChunks.Text -join "`n") | Should -Match 'Ada'

        $yamlPath = Join-Path $TestDrive 'source.yaml'
        Set-Content -Path $yamlPath -Value @(
            'service:'
            '  name: OfficeIMO'
            '  port: 443'
        ) -Encoding UTF8

        $yamlChunks = @(Get-OfficeDocumentChunk -Path $yamlPath)
        $yamlChunks.Count | Should -BeGreaterThan 0
        $yamlChunks[0].Kind.ToString() | Should -Be 'Yaml'
        ($yamlChunks.Text -join "`n") | Should -Match '\$\.service\.name'
        ($yamlChunks.Text -join "`n") | Should -Match 'OfficeIMO'
    }

    It 'reads Markdown tables and materializes deterministic table sidecars' {
        $path = Join-Path $TestDrive 'table.md'
        Set-Content -Path $path -Value @(
            '# Tables'
            ''
            '| Name | Score |'
            '| --- | ---: |'
            '| Ada | 42 |'
        ) -Encoding UTF8

        $tables = @(Get-OfficeDocumentTable -Path $path)
        $tables | Should -HaveCount 1
        $tables[0].Columns | Should -Be @('Name', 'Score')
        $tables[0].Rows[0][0] | Should -Be 'Ada'
        $tables[0].Rows[0][1] | Should -Be '42'

        $exports = @(Get-OfficeDocumentTable -Path $path -AsExport)
        $exports | Should -HaveCount 1
        $exports[0].Csv | Should -Match 'Name,Score'
        $exports[0].Markdown | Should -Match '\| Ada \| 42 \|'
        $exports[0].Json | Should -Match '"columns":\["Name","Score"\]'

        $outputDirectory = Join-Path $TestDrive 'table-exports'
        $materialized = @(Get-OfficeDocumentTable -Path $path -OutputDirectory $outputDirectory)
        $materialized | Should -HaveCount 3
        $materialized.Written | Should -Not -Contain $false
        Test-Path -LiteralPath (Join-Path $outputDirectory 'table-table-0000.csv') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $outputDirectory 'table-table-0000.md') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $outputDirectory 'table-table-0000.json') | Should -BeTrue
    }

    It 'reads folders using extension filters' {
        $folder = Join-Path $TestDrive 'reader-folder'
        New-Item -Path $folder -ItemType Directory | Out-Null
        Set-Content -Path (Join-Path $folder 'first.md') -Value '# First' -Encoding UTF8
        Set-Content -Path (Join-Path $folder 'skip.txt') -Value 'skip me' -Encoding UTF8

        $chunks = @(Get-OfficeDocumentChunk -FolderPath $folder -Extension md -NoRecurse)
        ($chunks.Location.Path | Select-Object -Unique) | Should -HaveCount 1
        ($chunks.Text -join "`n") | Should -Match 'First'
        ($chunks.Text -join "`n") | Should -Not -Match 'skip me'
    }

    It 'summarizes folder ingestion for adapter-backed formats' {
        $folder = Join-Path $TestDrive 'ingest-folder'
        New-Item -Path $folder -ItemType Directory | Out-Null
        Set-Content -Path (Join-Path $folder 'page.html') -Value '<html><body><p>Indexed HTML</p></body></html>' -Encoding UTF8
        Set-Content -Path (Join-Path $folder 'data.json') -Value '{"item":"Indexed JSON"}' -Encoding UTF8
        Set-Content -Path (Join-Path $folder 'skip.txt') -Value 'skip me' -Encoding UTF8

        Set-Content -Path (Join-Path $folder 'config.yaml') -Value "name: Indexed YAML" -Encoding UTF8

        $result = Get-OfficeDocumentIngest -FolderPath $folder -Extension html,json,yaml -NoRecurse
        $result.FilesScanned | Should -Be 3
        $result.FilesParsed | Should -Be 3
        $result.ChunksProduced | Should -Be 3
        ($result.Chunks.Text -join "`n") | Should -Match 'Indexed HTML'
        ($result.Chunks.Text -join "`n") | Should -Match 'Indexed JSON'
        ($result.Chunks.Text -join "`n") | Should -Match 'Indexed YAML'
        ($result.Chunks.Text -join "`n") | Should -Not -Match 'skip me'
    }
}
