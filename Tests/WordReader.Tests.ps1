BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force
}

Describe 'Word reader helpers' {
    It 'finds text and exposes bookmarks/fields' {
        $path = Join-Path $TestDrive 'WordReader.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph -Text 'Hello world'
        }

        $doc = Get-OfficeWord -Path $path
        try {
            $null = $doc.AddBookmark('Bookmark1')
            $paragraph = $doc.AddParagraph('Page')
            $null = $paragraph.AddField([OfficeIMO.Word.WordFieldType]::Page)
        } finally {
            Close-OfficeWord -Document $doc -Save
        }

        $matches = Find-OfficeWord -Path $path -Text 'Hello'
        $matches.Count | Should -BeGreaterThan 0

        $bookmarks = Get-OfficeWordBookmark -Path $path
        $bookmarks.Name | Should -Contain 'Bookmark1'

        $fields = Get-OfficeWordField -Path $path -FieldType Page
        $fields.Count | Should -BeGreaterThan 0
    }
}
