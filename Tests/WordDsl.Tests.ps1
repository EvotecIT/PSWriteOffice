BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force
}

Describe 'Word DSL surface' {
    It 'creates a document with canonical cmdlets' {
        $path = Join-Path $TestDrive 'DslCanonical.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordSection {
                Add-OfficeWordParagraph -Text 'Smoke test paragraph.'
                Add-OfficeWordList -Style 'Bulleted' {
                    Add-OfficeWordListItem -Text 'First'
                    Add-OfficeWordListItem -Text 'Second'
                }
            }
        }

        Test-Path $path | Should -BeTrue

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.Sections.Count | Should -BeGreaterThan 0
            $document.Paragraphs.Count | Should -BeGreaterThan 0
        } finally {
            $document.Dispose()
        }
    }

    It 'supports alias-style DSL with tables' {
        $path = Join-Path $TestDrive 'DslAlias.docx'
        $rows = @(
            [PSCustomObject]@{ Item = 'Alpha'; Qty = 1 }
            [PSCustomObject]@{ Item = 'Beta'; Qty = 5 }
        )

        New-OfficeWord -Path $path {
            WordSection {
                WordParagraph { WordText 'Alias smoke' }
                WordList {
                    WordListItem 'One'
                    WordListItem 'Two'
                }
                WordTable -Data $rows -Style 'GridTable1LightAccent1' {
                    WordTableCondition -FilterScript { $_.Qty -gt 2 } -BackgroundColor '#ffeeee'
                }
            }
        }

        Test-Path $path | Should -BeTrue
    }

    It 'supports reader helpers and save' {
        $path = Join-Path $TestDrive 'DslReaders.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'One'; Value = 1 }
        )

        New-OfficeWord -Path $path {
            Add-OfficeWordSection {
                Add-OfficeWordParagraph -Text 'Reader smoke.'
                Add-OfficeWordTable -InputObject $rows -Style 'TableGrid'
            }
        } | Out-Null

        $document = Get-OfficeWord -Path $path
        try {
            ($document | Get-OfficeWordSection).Count | Should -BeGreaterThan 0
            ($document | Get-OfficeWordParagraph).Count | Should -BeGreaterThan 0
            ($document | Get-OfficeWordTable).Count | Should -BeGreaterThan 0
            ($document | Get-OfficeWordParagraph | Select-Object -First 1 | Get-OfficeWordRun).Count | Should -BeGreaterThan 0

            $document | Save-OfficeWord | Out-Null
        } finally {
            Close-OfficeWord -Document $document
        }
    }

    It 'adds fields, watermarks, and protection' {
        $path = Join-Path $TestDrive 'DslProtected.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph {
                Add-OfficeWordText -Text 'Page '
                Add-OfficeWordField -Type Page
            }
            Add-OfficeWordWatermark -Text 'CONFIDENTIAL'
            Protect-OfficeWordDocument -Password 'secret'
        } | Out-Null

        Test-Path $path | Should -BeTrue
    }
}
