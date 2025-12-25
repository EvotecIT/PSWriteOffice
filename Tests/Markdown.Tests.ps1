BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force
}

Describe 'Markdown cmdlets' {
    It 'parses Markdown text into a document' {
        $doc = Get-OfficeMarkdown -Text "# Title`n`nHello"
        $doc | Should -BeOfType OfficeIMO.Markdown.MarkdownDoc
    }

    It 'converts objects to Markdown tables' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $markdown = $rows | ConvertTo-OfficeMarkdown
        $markdown | Should -Match 'Name'
        $markdown | Should -Match 'Value'
    }

    It 'converts Markdown to HTML' {
        $html = ConvertTo-OfficeMarkdownHtml -Text "# Title`n`nHello"
        $html | Should -Match '<h1'
        $html | Should -Match 'Title'
    }
}
