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
