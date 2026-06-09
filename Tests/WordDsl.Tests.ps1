BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')

    function Test-OfficeLoadedMethod {
        param(
            [Parameter(Mandatory)]
            [string] $TypeName,

            [Parameter(Mandatory)]
            [string] $MethodName
        )

        $type = [AppDomain]::CurrentDomain.GetAssemblies() |
            ForEach-Object { $_.GetType($TypeName, $false) } |
            Where-Object { $null -ne $_ } |
            Select-Object -First 1
        if ($null -eq $type) {
            throw "Unable to find loaded type '$TypeName'."
        }

        @($type.GetMethods() | Where-Object Name -eq $MethodName).Count -gt 0
    }
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

    It 'round-trips encrypted Word documents through lifecycle cmdlets' {
        if (-not (Test-OfficeLoadedMethod -TypeName 'OfficeIMO.Word.WordDocument' -MethodName 'LoadEncrypted')) {
            (Get-Command New-OfficeWord).Parameters.Keys | Should -Contain 'Password'
            (Get-Command Save-OfficeWord).Parameters.Keys | Should -Contain 'Password'
            (Get-Command Get-OfficeWord).Parameters.Keys | Should -Contain 'Password'
            return
        }

        $path = Join-Path $TestDrive 'EncryptedWord.docx'

        New-OfficeWord -Path $path -Password 'secret' {
            WordSection {
                WordParagraph -Text 'Encrypted Word value'
            }
        }

        { Get-ZipEntriesLocal -Path $path } | Should -Throw

        $document = Get-OfficeWord -Path $path -Password 'secret' -ReadOnly
        try {
            $document.Paragraphs.Text | Should -Contain 'Encrypted Word value'
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
                WordTable -InputObject $rows -Style 'GridTable1LightAccent1' {
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
            WordTable -InputObject $rows -Style 'GridTable1LightAccent1' {
                WordTableCell -Row 1 -Column 0 {
                    WordParagraph { WordText 'Checklist' }
                    WordImage -Path $imagePath -Width 24 -Height 24
                    WordList {
                        WordListItem 'Confirm issue coverage'
                        WordListItem 'Stage release notes'
                    }
                }

                WordTableCell -Row 1 -Column 1 {
                    WordTable -InputObject $nestedRows -NoHeader -Style 'TableGrid' {
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

    It 'adds same-paragraph Word breaks' {
        $path = Join-Path $TestDrive 'DslWordBreaks.docx'

        New-OfficeWord -Path $path {
            WordParagraph {
                WordText 'Line 1'
                WordBreak
                WordText 'Line 2'
            }
        } | Out-Null

        $documentXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'word/document.xml'
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($documentXml.NameTable)
        $namespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

        $documentXml.SelectNodes('//w:br', $namespaceManager).Count | Should -Be 1
        $documentXml.SelectSingleNode('//w:t[text()="Line 1"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $documentXml.SelectSingleNode('//w:t[text()="Line 2"]', $namespaceManager) | Should -Not -BeNullOrEmpty
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
            ($document | Get-OfficeWordParagraph | Select-Object -First 1 | Get-OfficeWordText).Count | Should -BeGreaterThan 0

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
            $document.Tables[0].LayoutMode.ToString() | Should -Be 'AutoFitToContents'
            $document.Tables[1].LayoutMode.ToString() | Should -Be 'AutoFitToWindow'
            $document.Tables[2].LayoutType.Value | Should -Be 'fixed'
        } finally {
            $document.Dispose()
        }
    }

    It 'adds Word pie charts inside the DSL' {
        $path = Join-Path $TestDrive 'DslWordPieChart.docx'
        $rows = @(
            [PSCustomObject]@{ Region = 'North America'; Revenue = 125000 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 98000 }
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 143000 }
        )

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph -Text 'Revenue mix'
            Add-OfficeWordChart -Type Pie -InputObject $rows -CategoryProperty Region -SeriesProperty Revenue -Title 'Regional Revenue Mix'
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.Charts.Count | Should -Be 1
            $document.Paragraphs.Where({ $_.IsChart }).Count | Should -Be 1
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
            Add-OfficeWordTable -InputObject $rows -View Transpose -Style TableGrid
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

    It 'adds Word line charts to an open document' {
        $path = Join-Path $TestDrive 'DslWordLineChart.docx'
        $rows = @(
            [PSCustomObject]@{ Month = 'Jan'; Sales = 10; Profit = 4 }
            [PSCustomObject]@{ Month = 'Feb'; Sales = 12; Profit = 5 }
            [PSCustomObject]@{ Month = 'Mar'; Sales = 15; Profit = 7 }
        )

        $document = New-OfficeWord -Path $path
        try {
            $chart = Add-OfficeWordChart -Document $document -Type Line -InputObject $rows -CategoryProperty Month -SeriesProperty Sales, Profit -Legend -XAxisTitle 'Month' -YAxisTitle 'Value' -Title 'Monthly Trend' -PassThru
            $chart.Title | Should -Be 'Monthly Trend'
            Save-OfficeWord -Document $document | Out-Null
        } finally {
            Close-OfficeWord -Document $document
        }

        $saved = Get-OfficeWord -Path $path -ReadOnly
        try {
            $saved.Charts.Count | Should -Be 1
            $saved.Paragraphs.Where({ $_.IsChart }).Count | Should -Be 1
        } finally {
            $saved.Dispose()
        }
    }

    It 'wraps OfficeIMO Word page setup, cover pages, equations, tab stops, and statistics' {
        $path = Join-Path $TestDrive 'DslWordRoadmap.docx'
        $omml = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x+1</m:t></m:r></m:oMath>'

        New-OfficeWord -Path $path {
            Add-OfficeWordCoverPage -Template Element -Abstract 'Executive summary' -CompanyEmail 'reports@example.test'
            Add-OfficeWordSection {
                Set-OfficeWordPageSetup -PageSize A4 -Orientation Landscape -Margin Narrow -Columns 2 -ColumnSpacing 720 -ColumnSeparator $true
                Add-OfficeWordParagraph {
                    Add-OfficeWordTabStop -Position 4320 -Alignment Decimal -Leader Dot
                    Add-OfficeWordText -Text "Score`t98.5"
                }
                Add-OfficeWordEquation -Omml $omml
            }
        }

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $document.CoverPage | Should -Not -BeNullOrEmpty
            $document.CoverPageProperties.Abstract | Should -Be 'Executive summary'
            $document.CoverPageProperties.CompanyEmail | Should -Be 'reports@example.test'
            $document.Sections[0].PageSettings.PageSize.ToString() | Should -Be 'A4'
            $document.Sections[0].PageOrientation.Value | Should -Be 'landscape'
            $document.Sections[0].Margins.Type.ToString() | Should -Be 'Narrow'
            $document.Sections[0].ColumnCount | Should -Be 2
            $document.Sections[0].ColumnsSpace | Should -Be 720
            $document.Sections[0].HasColumnSeparator | Should -BeTrue
            ($document.Paragraphs | Where-Object { $_.TabStops.Count -gt 0 }).Count | Should -BeGreaterThan 0
            $document.Equations.Count | Should -Be 1
        } finally {
            $document.Dispose()
        }

        $stats = Get-OfficeWordStatistics -Path $path
        $stats.Paragraphs | Should -BeGreaterThan 0
        $stats.Words | Should -BeGreaterThan 0
    }

    It 'joins Word documents through the thin append wrapper' {
        $basePath = Join-Path $TestDrive 'JoinBase.docx'
        $appendPath = Join-Path $TestDrive 'JoinAppendix.docx'
        $outputPath = Join-Path $TestDrive 'JoinMerged.docx'

        New-OfficeWord -Path $basePath {
            Add-OfficeWordParagraph -Text 'Base report'
        }
        New-OfficeWord -Path $appendPath {
            Add-OfficeWordParagraph -Text 'Appendix detail'
        }

        Join-OfficeWordDocument -Path $basePath -AppendPath $appendPath -OutputPath $outputPath

        $document = Get-OfficeWord -Path $outputPath -ReadOnly
        try {
            ($document.Paragraphs.Text -join "`n") | Should -Match 'Base report'
            ($document.Paragraphs.Text -join "`n") | Should -Match 'Appendix detail'
        } finally {
            $document.Dispose()
        }
    }

    It 'adds Word charts to paragraphs created inside table cells' {
        $path = Join-Path $TestDrive 'DslWordChartInCell.docx'
        $tableRows = @(
            [PSCustomObject]@{ Topic = 'Regional revenue'; Details = 'Pending' }
        )
        $chartRows = @(
            [PSCustomObject]@{ Region = 'North America'; Revenue = 125000 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 98000 }
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 143000 }
        )

        New-OfficeWord -Path $path {
            $table = Add-OfficeWordTable -InputObject $tableRows -Style 'GridTable1LightAccent1' -PassThru
            $paragraph = $table.Rows[1].Cells[1].AddParagraph()

            Add-OfficeWordChart -Paragraph $paragraph -Type Pie -InputObject $chartRows -CategoryProperty Region -SeriesProperty Revenue -Title 'Regional Revenue Mix'
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $cellCharts = $table.Rows[1].Cells[1].Paragraphs.Where({ $_.IsChart })

            $cellCharts.Count | Should -Be 1
            $cellCharts[0].Chart | Should -Not -BeNullOrEmpty
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

    It 'adds and reads footnotes and endnotes in the DSL' {
        $path = Join-Path $TestDrive 'DslNotes.docx'

        New-OfficeWord -Path $path {
            WordParagraph {
                WordText 'Availability is calculated from successful health probes'
                WordFootnote -Text 'Probe data excludes planned maintenance windows.'
            }

            WordParagraph {
                WordText 'The appendix keeps the full scoring details'
                WordEndnote -Text 'Scoring uses freshness, severity, and service ownership weights.'
            }
        } | Out-Null

        $footnotes = @(Get-OfficeWordFootnote -Path $path)
        $endnotes = @(Get-OfficeWordEndnote -Path $path)

        $footnotes.Count | Should -Be 1
        $endnotes.Count | Should -Be 1
        $footnotes[0].Text | Should -Match 'Probe data excludes planned maintenance windows'
        $endnotes[0].Text | Should -Match 'Scoring uses freshness'

        $entries = @(Get-ZipEntriesLocal -Path $path)
        $entries | Should -Contain 'word/footnotes.xml'
        $entries | Should -Contain 'word/endnotes.xml'
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

            Invoke-OfficeWordMailMerge -InputObject @{
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

    It 'wraps OfficeIMO Word image and shape editing helpers' {
        $path = Join-Path $TestDrive 'DslWordImagesAndShapes.docx'
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive

        New-OfficeWord -Path $path {
            WordParagraph {
                $image = WordImage -Path $imagePath -Width 32 -Height 32 -PassThru
                $image |
                    Set-OfficeWordImage -Width 48 -Height 24 -Title 'Status Logo' -Description 'Logo alt text' -HorizontalFlip $true -Rotation 5 -PassThru |
                    Should -Not -BeNullOrEmpty
            }

            WordParagraph {
                $shape = WordShape -Type RoundedRectangle -Width 144 -Height 36 -FillColor '#DDEEFF' -StrokeColor '#1F4E79' -StrokeWidth 1 -Title 'Callout' -Description 'Callout shape' -PassThru
                $shape |
                    Set-OfficeWordShape -Left 24 -Top 12 -Rotation 2 -ZIndex 3 -PassThru |
                    Should -Not -BeNullOrEmpty
                $shape.FillColorHex | Should -Be 'ddeeff'
                $shape.StrokeColorHex | Should -Be '1f4e79'
            }
        } | Out-Null

        $images = @(Get-OfficeWordImage -Path $path)
        $shapes = @(Get-OfficeWordShape -Path $path)

        $images.Count | Should -Be 1
        $images[0].Title | Should -Be 'Status Logo'
        $images[0].Description | Should -Be 'Logo alt text'
        $images[0].HorizontalFlip | Should -BeTrue

        $shapes.Count | Should -Be 1
        $shapes[0].FillColorHex | Should -Be 'ddeeff'
        $shapes[0].StrokeColorHex | Should -Be '1f4e79'
    }

    It 'wraps OfficeIMO Word table cell read and style helpers' {
        $path = Join-Path $TestDrive 'DslWordTableCells.docx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; State = 'Ready' }
            [PSCustomObject]@{ Name = 'Beta'; State = 'Blocked' }
        )

        New-OfficeWord -Path $path {
            $table = WordTable -InputObject $rows -Style TableGrid -PassThru
            $cell = $table |
                Get-OfficeWordTableCell -Row 1 -Column 1 |
                Set-OfficeWordTableCell -ShadingFillColor '#DDEEFF' -Width 2400 -WidthType Dxa -WrapText $false -FitText $true -PassThru

            $cell.ShadingFillColorHex | Should -Be 'DDEEFF'
            $cell.Width | Should -Be 2400
            $cell.FitText | Should -BeTrue
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $table = $document.Tables[0]
            $cell = $table | Get-OfficeWordTableCell -Row 1 -Column 1

            $cell.ShadingFillColorHex | Should -Be 'DDEEFF'
            $cell.Width | Should -Be 2400
            $cell.WidthType.Value | Should -Be 'dxa'
        } finally {
            $document.Dispose()
        }
    }

    It 'wraps OfficeIMO Word paragraph and text style helpers' {
        $path = Join-Path $TestDrive 'DslWordStyles.docx'

        New-OfficeWord -Path $path {
            $paragraph = WordParagraph -Text 'Executive Summary' -PassThru
            $paragraph |
                Set-OfficeWordParagraphStyle -Style Heading2 -Alignment Center -SpacingBeforePoints 6 -SpacingAfterPoints 12 -IndentationBeforePoints 18 -KeepWithNext $true -PassThru |
                Should -Not -BeNullOrEmpty

            $textParagraph = WordParagraph -Text 'Initial' -PassThru
            $textItem = @($textParagraph.GetRuns())[0]
            $textItem |
                Set-OfficeWordTextStyle -Text 'Styled text' -Bold $true -Italic $true -Underline Single -Color '#C00000' -FontSize 14 -FontFamily 'Aptos' -Highlight Yellow -CapsStyle SmallCaps -Strike $true -Style Heading2Char -PassThru |
                Should -Not -BeNullOrEmpty
        } | Out-Null

        $document = Get-OfficeWord -Path $path -ReadOnly
        try {
            $styledParagraph = $document.Paragraphs | Where-Object Text -EQ 'Executive Summary' | Select-Object -First 1
            $styledParagraph.Style.ToString() | Should -Be 'Heading2'
            $styledParagraph.ParagraphAlignment.Value | Should -Be 'center'
            $styledParagraph.LineSpacingBeforePoints | Should -Be 6
            $styledParagraph.LineSpacingAfterPoints | Should -Be 12
            $styledParagraph.IndentationBeforePoints | Should -Be 18
            $styledParagraph.KeepWithNext | Should -BeTrue

            $styledText = $document.Paragraphs |
                Where-Object Text -EQ 'Styled text' |
                ForEach-Object { $_.GetRuns() } |
                Select-Object -First 1

            $styledText.Text | Should -Be 'Styled text'
            $styledText.Bold | Should -BeTrue
            $styledText.Italic | Should -BeTrue
            $styledText.Underline.Value | Should -Be 'single'
            $styledText.ColorHex | Should -Be 'c00000'
            $styledText.FontSize | Should -Be 14
            $styledText.FontFamily | Should -Be 'Aptos'
            $styledText.Highlight.Value | Should -Be 'yellow'
            $styledText.CapsStyle.ToString() | Should -Be 'SmallCaps'
            $styledText.Strike | Should -BeTrue
            $styledText.CharacterStyle.ToString() | Should -Be 'Heading2Char'
        } finally {
            $document.Dispose()
        }
    }
}
