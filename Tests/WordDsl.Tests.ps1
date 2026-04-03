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

    It 'supports paragraphs, lists, images, and nested tables inside table cells' {
        $path = Join-Path $TestDrive 'DslTableCells.docx'
        $imagePath = Join-Path $TestDrive 'CellImage.png'
        $fixturePath = Join-Path $PSScriptRoot 'Assets/CellImage.png'
        $rows = @(
            [PSCustomObject]@{
                Topic   = 'Release readiness'
                Details = 'Pending'
            }
        )

        $nestedRows = @(
            [PSCustomObject]@{ Step = 'Validate'; State = 'Ready' }
        )

        Copy-Item -LiteralPath $fixturePath -Destination $imagePath -Force

        New-OfficeWord -Path $path {
            WordTable -Data $rows -Style 'GridTable1LightAccent1' {
                WordTableCell -Row 1 -Column 0 {
                    WordParagraph { WordText 'Checklist' }
                    WordImage -Path $imagePath -Width 24 -Height 24
                    WordList {
                        WordListItem 'Confirm issue coverage'
                        WordListItem 'Stage release notes'
                    }
                }

                WordTableCell -Row 1 -Column 1 {
                    WordTable -Data $nestedRows -SkipHeader -Style 'TableGrid'
                }
            }
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $cellTexts = $table.Rows[1].Cells[0].Paragraphs |
                ForEach-Object Text |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $listTexts = $document.Lists |
                ForEach-Object ListItems |
                ForEach-Object Text |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

            $cellTexts | Should -Contain 'Checklist'
            $cellTexts | Should -Contain 'Confirm issue coverage'
            $listTexts | Should -Contain 'Confirm issue coverage'
            $listTexts | Should -Contain 'Stage release notes'
            ($table.Rows[1].Cells[0].Paragraphs | Where-Object IsImage).Count | Should -Be 1

            $table.HasNestedTables | Should -BeTrue
            $table.NestedTables.Count | Should -Be 1
            $table.NestedTables[0].Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Validate'
            $table.NestedTables[0].Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'Ready'
        } finally {
            $document.Dispose()
        }
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

    It 'supports hyperlinks and document properties in the DSL' {
        $path = Join-Path $TestDrive 'DslLinksAndProperties.docx'

        New-OfficeWord -Path $path {
            Set-OfficeWordDocumentProperty -Name Title -Value 'DSL document'
            Set-OfficeWordDocumentProperty -Name Creator -Value 'PSWriteOffice'
            Set-OfficeWordDocumentProperty -Name BuildNumber -Value 21 -Custom

            WordParagraph {
                WordText 'Open '
                WordHyperlink -Text 'Example' -Url 'https://example.org' -Styled -Tooltip 'External link'
                WordText ' or jump to '
                WordHyperlink -Text 'Summary' -Anchor 'Summary'
            }

            WordParagraph {
                WordText 'Summary'
                WordBookmark -Name 'Summary'
            }
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.HyperLinks.Count | Should -Be 2
            $document.BuiltinDocumentProperties.Title | Should -Be 'DSL document'
            $document.BuiltinDocumentProperties.Creator | Should -Be 'PSWriteOffice'
            $document.CustomDocumentProperties['BuildNumber'].Value | Should -Be 21
        } finally {
            $document.Dispose()
        }
    }

    It 'supports background colors and mail merge in the DSL' {
        $path = Join-Path $TestDrive 'DslBackgroundMailMerge.docx'

        New-OfficeWord -Path $path {
            Set-OfficeWordBackground -Color '#ddeeff'

            Add-OfficeWordParagraph {
                Add-OfficeWordText -Text 'Dear '
                Add-OfficeWordField -Type MergeField -Parameters '"FirstName"'
                Add-OfficeWordText -Text ','
            }

            Add-OfficeWordParagraph {
                Add-OfficeWordText -Text 'Order '
                Add-OfficeWordField -Type MergeField -Parameters '"OrderId"'
                Add-OfficeWordText -Text ' is ready.'
            }

            Invoke-OfficeWordMailMerge -Data @{
                FirstName = 'Jane'
                OrderId   = 77
            }
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.Background.Color | Should -Be 'ddeeff'
            $document.Fields.Where({ $_.FieldType -eq [OfficeIMO.Word.WordFieldType]::MergeField }).Count | Should -Be 0
        } finally {
            $document.Dispose()
        }

        (Find-OfficeWord -Path $path -Text 'Jane').Count | Should -BeGreaterThan 0
        (Find-OfficeWord -Path $path -Text '77').Count | Should -BeGreaterThan 0
    }
}
