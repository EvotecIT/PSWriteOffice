BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'Reader cmdlets' {
    It 'exposes built-in and modular Reader capabilities' {
        $capabilities = Get-OfficeDocumentCapability

        $capabilities.Id | Should -Contain 'officeimo.reader.word'
        $capabilities.Id | Should -Contain 'officeimo.reader.excel'
        $capabilities.Id | Should -Contain 'officeimo.reader.powerpoint'
        $capabilities.Id | Should -Contain 'officeimo.reader.pdf'
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
}
