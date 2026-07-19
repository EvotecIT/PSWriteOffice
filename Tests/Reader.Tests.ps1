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
        $capabilities.Id | Should -Contain 'officeimo.reader.markdown'
        $capabilities.Id | Should -Contain 'officeimo.reader.email'
        $capabilities.Id | Should -Contain 'officeimo.reader.email.store'
        $capabilities.Id | Should -Contain 'officeimo.reader.email.address-book'
        $capabilities.Id | Should -Contain 'officeimo.reader.pdf'
        $capabilities.Id | Should -Contain 'officeimo.reader.html'
        $capabilities.Id | Should -Contain 'officeimo.reader.csv'
        $capabilities.Id | Should -Contain 'officeimo.reader.json'
        $capabilities.Id | Should -Contain 'officeimo.reader.xml'
        $capabilities.Id | Should -Contain 'officeimo.reader.yaml'
        $capabilities.Id | Should -Contain 'officeimo.reader.zip'
        $capabilities.Id | Should -Contain 'officeimo.reader.epub'
        $capabilities.Id | Should -Contain 'officeimo.reader.visio'
        $capabilities.Id | Should -Contain 'officeimo.reader.rtf'
    }

    It 'exports Reader table, visual, asset, and ingestion cmdlets' {
        Get-Command -Name Get-OfficeDocumentTable -ErrorAction Stop | Should -Not -BeNullOrEmpty
        Get-Command -Name Get-OfficeDocumentVisual -ErrorAction Stop | Should -Not -BeNullOrEmpty
        Get-Command -Name Get-OfficeDocumentAsset -ErrorAction Stop | Should -Not -BeNullOrEmpty
        Get-Command -Name Get-OfficeDocumentIngest -ErrorAction Stop | Should -Not -BeNullOrEmpty
        Get-Command -Name Read-OfficeDocumentTable -ErrorAction Stop | Should -Not -BeNullOrEmpty
        Get-Command -Name Read-OfficeDocumentVisual -ErrorAction Stop | Should -Not -BeNullOrEmpty
        Get-Command -Name Read-OfficeDocumentAsset -ErrorAction Stop | Should -Not -BeNullOrEmpty
    }

    It 'accepts a caller-configured immutable Reader' {
        $handlerId = 'pswriteoffice.test.custom'
        $registrationType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader.Core' -TypeName 'OfficeIMO.Reader.ReaderHandlerRegistration' -CommandName 'Get-OfficeDocumentCapability'
        $builderType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader.Core' -TypeName 'OfficeIMO.Reader.OfficeDocumentReaderBuilder' -CommandName 'Get-OfficeDocumentCapability'
        $inputKindType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader.Core' -TypeName 'OfficeIMO.Reader.ReaderInputKind' -CommandName 'Get-OfficeDocumentCapability'
        $chunkType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader.Core' -TypeName 'OfficeIMO.Reader.ReaderChunk' -CommandName 'Get-OfficeDocumentCapability'

        $registration = [Activator]::CreateInstance($registrationType)
        $registration.Id = $handlerId
        $registration.DisplayName = 'Test custom Reader'
        $registration.Kind = [Enum]::Parse($inputKindType, 'Text')
        $registration.Extensions = [string[]]@('.custom')
        $readPathType = $registrationType.GetProperty('ReadPath').PropertyType
        $registration.ReadPath = [System.Management.Automation.LanguagePrimitives]::ConvertTo({
            param($Path, $Options, $CancellationToken)

            [Array]::CreateInstance($chunkType, 0)
        }.GetNewClosure(), $readPathType)

        $builder = [Activator]::CreateInstance($builderType)
        $builderType.GetMethod('AddHandler').Invoke($builder, @($registration, $false)) | Out-Null
        $reader = $builderType.GetMethod('Build').Invoke($builder, @())

        $capabilities = @(Get-OfficeDocumentCapability -Reader $reader -ExcludeBuiltIn)
        $customCapability = $capabilities | Where-Object Id -EQ $handlerId
        $customCapability | Should -HaveCount 1
        $customCapability.Extensions | Should -Contain '.custom'
        (Get-Command Get-OfficeDocumentChunk).Parameters.Keys | Should -Contain 'Reader'
        (Get-Command New-OfficeDocumentReader).Parameters.Keys | Should -Contain 'ReaderAllOptions'

        $customPath = Join-Path $TestDrive 'source.custom'
        Set-Content -Path $customPath -Value 'custom reader input' -Encoding UTF8

        { Get-OfficeDocumentChunk -Path $customPath -Reader $reader -NoMarkdownHeadingChunks } |
            Should -Throw -ExpectedMessage '*immutable OfficeIMO 3 reader*'
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

    It 'reads RTF files through the semantic Reader adapter' {
        $path = Join-Path $TestDrive 'source.rtf'
        New-OfficeRtf -Path $path -Text 'Reader RTF adapter', 'Semantic chunk text' | Out-Null

        $chunks = @(Get-OfficeDocumentChunk -Path $path)
        $chunks.Count | Should -BeGreaterThan 0
        $chunks[0].Kind.ToString() | Should -Be 'Rtf'
        ($chunks.Text -join "`n") | Should -Match 'Reader RTF adapter'
        ($chunks.Text -join "`n") | Should -Match 'Semantic chunk text'
        ($chunks.Text -join "`n") | Should -Not -Match '\\rtf1'
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

    It 'reads and materializes embedded document assets' {
        $path = Join-Path $TestDrive 'assets.docx'
        $imagePath = Join-Path $PSScriptRoot 'Assets\CellImage.png'

        New-OfficeWord -Path $path {
            WordSection {
                WordParagraph { WordText 'Document with embedded asset' }
                WordParagraph { WordImage -Path $imagePath -Width 24 -Height 24 }
            }
        } | Out-Null

        $assets = @(Get-OfficeDocumentAsset -Path $path -Kind image)
        $assets.Count | Should -BeGreaterThan 0
        $assets[0].Kind | Should -Be 'image'
        $assets[0].MediaType | Should -Match 'image'
        $assets[0].PayloadBytes.Length | Should -BeGreaterThan 0

        $outputDirectory = Join-Path $TestDrive 'asset-exports'
        $materialized = @(Get-OfficeDocumentAsset -Path $path -Kind image -OutputDirectory $outputDirectory -ValidatePayloadHash)
        ($materialized | Where-Object Written).Count | Should -BeGreaterThan 0
        Test-Path -LiteralPath $materialized[0].Path | Should -BeTrue
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
