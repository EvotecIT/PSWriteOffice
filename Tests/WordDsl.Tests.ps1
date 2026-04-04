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
                    WordTable -Data $nestedRows -SkipHeader -Style 'TableGrid' {
                        WordTableCell -Row 0 -Column 0 {
                            WordParagraph { WordText 'Nested detail' }
                        }
                    }
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
            $nestedCellTexts = $table.NestedTables[0].Rows[0].Cells[0].Paragraphs |
                ForEach-Object Text |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $outerDetailTexts = $table.Rows[1].Cells[1].Paragraphs |
                ForEach-Object Text |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

            $cellTexts | Should -Contain 'Checklist'
            $cellTexts | Should -Contain 'Confirm issue coverage'
            $listTexts | Should -Contain 'Confirm issue coverage'
            $listTexts | Should -Contain 'Stage release notes'
            ($table.Rows[1].Cells[0].Paragraphs | Where-Object IsImage).Count | Should -Be 1

            $table.HasNestedTables | Should -BeTrue
            $table.NestedTables.Count | Should -Be 1
            $nestedCellTexts | Should -Contain 'Nested detail'
            $outerDetailTexts | Should -Not -Contain 'Nested detail'
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

    It 'supports explicit Word table layout modes' {
        $path = Join-Path $TestDrive 'DslTableLayouts.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'One'; Value = 1 }
            [PSCustomObject]@{ Name = 'Two'; Value = 2 }
        )

        New-OfficeWord -Path $path {
            Add-OfficeWordTable -InputObject $rows -Style TableGrid -Layout AutoFitToContents
            Add-OfficeWordTable -InputObject $rows -Style TableGrid -Layout AutoFitToWindow
            Add-OfficeWordTable -InputObject $rows -Style TableGrid -Layout Fixed
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.Tables.Count | Should -Be 3
            $document.Tables[0].LayoutMode | Should -Be ([OfficeIMO.Word.WordTableLayoutType]::AutoFitToContents)
            $document.Tables[1].LayoutMode | Should -Be ([OfficeIMO.Word.WordTableLayoutType]::AutoFitToWindow)
            $document.Tables[2].LayoutType | Should -Be ([DocumentFormat.OpenXml.Wordprocessing.TableLayoutValues]::Fixed)
        } finally {
            $document.Dispose()
        }
    }

    It 'supports transposed Word tables' {
        $path = Join-Path $TestDrive 'DslTableTranspose.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'One'; Value = 1 }
            [PSCustomObject]@{ Name = 'Two'; Value = 2 }
        )

        New-OfficeWord -Path $path {
            Add-OfficeWordTable -InputObject $rows -Transpose -Style TableGrid
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $table.RowsCount | Should -Be 3
            $table.Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Property'
            $table.Rows[0].Cells[1].Paragraphs[0].Text | Should -Be 'Row1'
            $table.Rows[0].Cells[2].Paragraphs[0].Text | Should -Be 'Row2'
            $table.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Name'
            $table.Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'One'
            $table.Rows[1].Cells[2].Paragraphs[0].Text | Should -Be 'Two'
            $table.Rows[2].Cells[0].Paragraphs[0].Text | Should -Be 'Value'
            $table.Rows[2].Cells[1].Paragraphs[0].Text | Should -Be '1'
            $table.Rows[2].Cells[2].Paragraphs[0].Text | Should -Be '2'
        } finally {
            $document.Dispose()
        }
    }

    It 'ignores null rows when exporting Word tables' {
        $path = Join-Path $TestDrive 'DslTableNullRows.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'One'; Value = 1 }
            $null
            [PSCustomObject]@{ Name = 'Two'; Value = 2 }
        )

        New-OfficeWord -Path $path {
            Add-OfficeWordTable -InputObject $rows -Style TableGrid
        } | Out-Null

        Test-Path $path | Should -BeTrue

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.Tables.Count | Should -Be 1
            $table = $document.Tables[0]
            $table.RowsCount | Should -Be 3
            $table.Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'One'
            $table.Rows[2].Cells[0].Paragraphs[0].Text | Should -Be 'Two'
        } finally {
            $document.Dispose()
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

    It 'updates Word text and hyperlink metadata from a path' {
        $path = Join-Path $TestDrive 'DslReplaceText.docx'

        New-OfficeWord -Path $path {
            WordParagraph {
                WordText 'FY24 status'
                WordHyperlink -Text 'Portal FY24' -Url 'https://old.example.com/FY24' -Tooltip 'FY24 link'
            }

            WordParagraph {
                WordHyperlink -Text 'Jump FY24' -Anchor 'FY24Summary' -Tooltip 'FY24 anchor'
            }

            WordParagraph {
                WordText 'Summary'
                WordBookmark -Name 'FY24Summary'
            }
        } | Out-Null

        $replacements = Update-OfficeWordText -Path $path -OldValue 'FY24' -NewValue 'FY25' -IncludeHyperlinkText -IncludeHyperlinkUri -IncludeHyperlinkAnchor -IncludeHyperlinkTooltip
        $replacements | Should -BeGreaterThan 0

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            (Find-OfficeWord -Document $document -Text 'FY25').Count | Should -BeGreaterThan 0
            $document.HyperLinks[0].Text | Should -Be 'Portal FY25'
            $document.HyperLinks[0].Uri.OriginalString | Should -Be 'https://old.example.com/FY25'
            $document.HyperLinks[0].Tooltip | Should -Be 'FY25 link'
            $document.HyperLinks[1].Text | Should -Be 'Jump FY25'
            $document.HyperLinks[1].Anchor | Should -Be 'FY25Summary'
            $document.HyperLinks[1].Tooltip | Should -Be 'FY25 anchor'
            $document.Bookmarks.Name | Should -Contain 'FY25Summary'
            $document.Bookmarks.Name | Should -Not -Contain 'FY24Summary'
        } finally {
            $document.Dispose()
        }
    }

    It 'updates relative hyperlink targets when requested' {
        $path = Join-Path $TestDrive 'DslReplaceRelativeLink.docx'

        New-OfficeWord -Path $path {
            WordParagraph {
                WordHyperlink -Text 'Relative FY24' -Url 'https://placeholder.invalid/FY24'
            }
        } | Out-Null

        $editable = Get-OfficeWord -Path $path
        try {
            $editable.HyperLinks[0].Uri = [Uri]::new('../reports/FY24.docx', [UriKind]::RelativeOrAbsolute)
            Save-OfficeWord -Document $editable | Out-Null
        } finally {
            $editable.Dispose()
        }

        $replacements = Update-OfficeWordText -Path $path -OldValue 'FY24' -NewValue 'FY25' -IncludeHyperlinkUri
        $replacements | Should -Be 1

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.HyperLinks[0].Uri.OriginalString | Should -Be '../reports/FY25.docx'
        } finally {
            $document.Dispose()
        }
    }

    It 'tracks current Word documents for update and close operations' {
        $pathOne = Join-Path $TestDrive 'TrackedOne.docx'
        $pathTwo = Join-Path $TestDrive 'TrackedTwo.docx'

        $docOne = New-OfficeWord -Path $pathOne
        $docTwo = New-OfficeWord -Path $pathTwo

        try {
            $docOne.AddParagraph('First tracked document') | Out-Null
            $docTwo.AddParagraph('Second tracked FY24 document') | Out-Null

            Update-OfficeWordText -OldValue 'FY24' -NewValue 'FY25' | Should -Be 1

            Close-OfficeWord -Save
            Close-OfficeWord -All -Save
        } finally {
            foreach ($doc in @($docOne, $docTwo)) {
                if ($null -ne $doc) {
                    try {
                        $doc.Dispose()
                    } catch {
                    }
                }
            }
        }

        Test-Path $pathOne | Should -BeTrue
        Test-Path $pathTwo | Should -BeTrue

        $savedOne = Get-OfficeWord -Path $pathOne -ReadOnly
        $savedTwo = Get-OfficeWord -Path $pathTwo -ReadOnly
        try {
            (Find-OfficeWord -Document $savedOne -Text 'First tracked document').Count | Should -Be 1
            (Find-OfficeWord -Document $savedTwo -Text 'FY25').Count | Should -Be 1
        } finally {
            $savedOne.Dispose()
            $savedTwo.Dispose()
        }
    }

    It 'does not fall back to the tracked document when -Document is null' {
        $path = Join-Path $TestDrive 'NullDocumentGuard.docx'
        $doc = New-OfficeWord -Path $path

        try {
            $nullDocument = $null
            { Close-OfficeWord -Document $nullDocument } | Should -Throw

            { $doc.AddParagraph('Still tracked after null guard') | Out-Null } | Should -Not -Throw
            Close-OfficeWord -Document $doc -Save
        } finally {
            if ($null -ne $doc) {
                try {
                    $doc.Dispose()
                } catch {
                }
            }
        }

        $saved = Get-OfficeWord -Path $path -ReadOnly
        try {
            (Find-OfficeWord -Document $saved -Text 'Still tracked after null guard').Count | Should -Be 1
        } finally {
            $saved.Dispose()
        }
    }
}
