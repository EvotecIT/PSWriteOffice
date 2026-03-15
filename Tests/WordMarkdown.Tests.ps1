BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Force -Global
}

Describe 'Word Markdown conversions' {
    It 'converts Word to Markdown text and file' {
        $docPath = Join-Path $TestDrive 'MarkdownSource.docx'
        $markdownPath = Join-Path $TestDrive 'MarkdownSource.md'

        New-OfficeWord -Path $docPath {
            Add-OfficeWordParagraph -Text 'Quarterly Report' -Style Heading1
            Add-OfficeWordParagraph -Text 'Hello Markdown'
        } | Out-Null

        $markdown = ConvertTo-OfficeWordMarkdown -Path $docPath
        $markdown | Should -Match '# Quarterly Report'
        $markdown | Should -Match 'Hello Markdown'

        $file = ConvertTo-OfficeWordMarkdown -Path $docPath -OutputPath $markdownPath -PassThru
        $file | Should -BeOfType System.IO.FileInfo
        (Get-Content -Path $markdownPath -Raw) | Should -Match 'Quarterly Report'
    }

    It 'converts Markdown text and Markdown documents to Word' {
        $docPath = Join-Path $TestDrive 'MarkdownRoundtrip.docx'
        $markdown = "# Title`n`nBody text"

        $file = ConvertFrom-OfficeWordMarkdown -Markdown $markdown -OutputPath $docPath -PassThru
        $file | Should -BeOfType System.IO.FileInfo
        Test-Path $docPath | Should -BeTrue

        $paragraphs = Get-OfficeWordParagraph -Path $docPath
        ($paragraphs.Text -join "`n") | Should -Match 'Title'
        ($paragraphs.Text -join "`n") | Should -Match 'Body text'

        $markdownDocument = Get-OfficeMarkdown -Text "## Pipeline`n`nGenerated"
        $document = $markdownDocument | ConvertFrom-OfficeWordMarkdown
        try {
            $document.Paragraphs.Count | Should -BeGreaterThan 0
        } finally {
            $document.Dispose()
        }
    }
}
