BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Force -Global
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

    It 'adds content controls and table of contents' {
        $path = Join-Path $TestDrive 'DslContentControls.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph { Add-OfficeWordContentControl -Text 'Client' -Alias 'ClientName' }
            Add-OfficeWordParagraph { Add-OfficeWordCheckBox -Checked -Alias 'Approved' }
            Add-OfficeWordParagraph { Add-OfficeWordDatePicker -Date (Get-Date) -Alias 'DueDate' }
            Add-OfficeWordParagraph { Add-OfficeWordDropDownList -Items 'Low','Medium','High' -Alias 'Priority' }
            Add-OfficeWordParagraph { Add-OfficeWordComboBox -Items 'Red','Blue' -DefaultValue 'Blue' -Alias 'Color' }
            Add-OfficeWordParagraph { Add-OfficeWordRepeatingSection -SectionTitle 'Items' -Alias 'LineItems' }
            Add-OfficeWordTableOfContent -Style Template1
            Update-OfficeWordFields
        } | Out-Null

        Test-Path $path | Should -BeTrue

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.StructuredDocumentTags.Count | Should -BeGreaterThan 0
            $document.CheckBoxes.Count | Should -BeGreaterThan 0
            $document.DatePickers.Count | Should -BeGreaterThan 0
            $document.DropDownLists.Count | Should -BeGreaterThan 0
            $document.ComboBoxes.Count | Should -BeGreaterThan 0
            $document.RepeatingSections.Count | Should -BeGreaterThan 0
            $document.TableOfContent | Should -Not -BeNullOrEmpty
        } finally {
            $document.Dispose()
        }
    }
}
