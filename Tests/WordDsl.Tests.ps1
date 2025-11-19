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
}
