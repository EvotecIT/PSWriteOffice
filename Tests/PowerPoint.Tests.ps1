BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Force -Global

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'PowerPoint cmdlets' {
    It 'creates a presentation with shapes, tables, media, and notes' {
        $path = Join-Path $TestDrive 'PowerPointContent.pptx'
        $presentation = New-OfficePowerPoint -FilePath $path
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive

        $layouts = Get-OfficePowerPointLayout -Presentation $presentation
        $layouts.Count | Should -BeGreaterThan 0
        $layoutType = $layouts | Where-Object { $_.Type } | Select-Object -First 1
        if ($layoutType) {
            $layoutMaster = $layoutType.MasterIndex
            $layoutIndex = $layoutType.LayoutIndex
            $slide2 = Add-OfficePowerPointSlide -Presentation $presentation -LayoutType $layoutType.Type -Master $layoutType.MasterIndex
        } elseif ($layouts[0].Name) {
            $layoutMaster = $layouts[0].MasterIndex
            $layoutIndex = $layouts[0].LayoutIndex
            $slide2 = Add-OfficePowerPointSlide -Presentation $presentation -LayoutName $layouts[0].Name -Master $layouts[0].MasterIndex
        } else {
            $layoutMaster = $layouts[0].MasterIndex
            $layoutIndex = $layouts[0].LayoutIndex
            $slide2 = Add-OfficePowerPointSlide -Presentation $presentation -Layout $layouts[0].LayoutIndex -Master $layouts[0].MasterIndex
        }

        $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
        Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Status Update'
        $shape = Add-OfficePowerPointShape -Slide $slide -ShapeType Rectangle -X 40 -Y 40 -Width 200 -Height 80 -FillColor '#DDEEFF' -OutlineColor '#1F4E79' -OutlineWidth 1

        $rows = @(
            [PSCustomObject]@{ Item = 'Alpha'; Qty = 10 }
            [PSCustomObject]@{ Item = 'Beta'; Qty = 20 }
        )
        $table = Add-OfficePowerPointTable -Slide $slide -Data $rows -X 40 -Y 140 -Width 360 -Height 200
        $image = Add-OfficePowerPointImage -Slide $slide -Path $imagePath -X 420 -Y 40 -Width 120 -Height 90
        $bullets = Add-OfficePowerPointBullets -Slide $slide -Bullets 'Wins','Risks','Next Steps' -X 420 -Y 150 -Width 250 -Height 200
        Set-OfficePowerPointNotes -Slide $slide -Text 'Keep this under five minutes.'

        $slide.Shapes.Count | Should -BeGreaterThan 0
        @($slide.Tables).Count | Should -BeGreaterThan 0
        $table.Rows | Should -BeGreaterThan 0
        $shape | Should -Not -BeNullOrEmpty
        $slide.Pictures.Count | Should -BeGreaterThan 0
        $image | Should -Not -BeNullOrEmpty
        $bullets.Paragraphs.Count | Should -Be 3
        $bullets.Paragraphs[0].Text | Should -Be 'Wins'
        $notes = Get-OfficePowerPointNotes -Slide $slide
        $notes.Text | Should -Be 'Keep this under five minutes.'
        $notes.HasNotes | Should -BeTrue
        $slideSummary = Get-OfficePowerPointSlideSummary -Slide $slide
        $slideSummary.Title | Should -Be 'Status Update'
        $slideSummary.HasNotes | Should -BeTrue
        $slideSummary.NotesText | Should -Be 'Keep this under five minutes.'
        $slideSummary.ShapeCount | Should -BeGreaterThan 0
        $slideSummary.PictureCount | Should -Be 1
        $slideSummary.TableCount | Should -Be 1
        $slideSummary.PlaceholderCount | Should -BeGreaterThan 0
        $slideSummary.LayoutPlaceholderCount | Should -BeGreaterThan 0
        $shapeInfo = Get-OfficePowerPointShape -Slide $slide
        $shapeInfo.Count | Should -BeGreaterThan 0
        $shapeInfo | Where-Object Kind -eq 'Picture' | Should -HaveCount 1
        $shapeInfo | Where-Object Kind -eq 'Table' | Should -HaveCount 1
        $shapeInfo | Where-Object Kind -eq 'AutoShape' | Should -HaveCount 1
        ($shapeInfo | Where-Object Kind -eq 'Picture' | Select-Object -First 1).MimeType | Should -Be 'image/bmp'
        ($shapeInfo | Where-Object Kind -eq 'Table' | Select-Object -First 1).RowCount | Should -Be 3
        ($shapeInfo | Where-Object Kind -eq 'Table' | Select-Object -First 1).ColumnCount | Should -Be 2
        $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType Title
        $placeholder.Text | Should -Be 'Status Update'
        $placeholderUpdate = Set-OfficePowerPointPlaceholderText -Slide $slide -PlaceholderType Title -Text 'Status Update v2' -PassThru
        $placeholderUpdate.Text | Should -Be 'Status Update v2'
        $layoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder -Slide $slide
        $layoutPlaceholders.Count | Should -BeGreaterThan 0

        $slide2 | Should -Not -BeNullOrEmpty

        Save-OfficePowerPoint -Presentation $presentation

        Test-Path $path | Should -BeTrue
        $presentation.Dispose()
        $presentation = $null

        $reloaded = Get-OfficePowerPoint -FilePath $path
        try {
            $reloaded.Slides.Count | Should -Be 2

            $reloadedSlide = Get-OfficePowerPointSlide -Presentation $reloaded -Index 1
            $reloadedPlaceholder = Get-OfficePowerPointPlaceholder -Slide $reloadedSlide -PlaceholderType Title
            if (-not $reloadedPlaceholder) {
                $reloadedPlaceholder = Get-OfficePowerPointPlaceholder -Slide $reloadedSlide -PlaceholderType CenteredTitle
            }
            $reloadedPlaceholder.Text | Should -Be 'Status Update v2'
            $reloadedNotes = Get-OfficePowerPointNotes -Slide $reloadedSlide
            $reloadedNotes.Text | Should -Be 'Keep this under five minutes.'
            $reloadedNotes.SlideIndex | Should -Be 1
            $reloadedSummary = Get-OfficePowerPointSlideSummary -Presentation $reloaded -Index 1
            $reloadedSummary.Title | Should -Be 'Status Update v2'
            $reloadedSummary.SlideIndex | Should -Be 1
            $reloadedSummary.PictureCount | Should -Be 1
            $reloadedSummary.TableCount | Should -Be 1
            $reloadedShapeInfo = Get-OfficePowerPointShape -Presentation $reloaded -Index 1
            $reloadedShapeInfo.Count | Should -BeGreaterThan 0
            ($reloadedShapeInfo | Where-Object Kind -eq 'Picture').Count | Should -Be 1
            ($reloadedShapeInfo | Where-Object Kind -eq 'Table').Count | Should -Be 1
            ($reloadedShapeInfo | Where-Object Kind -eq 'Picture' | Select-Object -First 1).SlideIndex | Should -Be 1
            @($reloadedSlide.Tables).Count | Should -BeGreaterThan 0
            $reloadedSlide.Pictures.Count | Should -BeGreaterThan 0
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'reads notes without creating empty notes parts' {
        $path = Join-Path $TestDrive 'PowerPointNotesRead.pptx'
        $presentation = New-OfficePowerPoint -FilePath $path

        $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
        Set-OfficePowerPointSlideTitle -Slide $slide -Title 'No notes yet'

        $notes = Get-OfficePowerPointNotes -Slide $slide -IncludeEmpty
        $notes.HasNotes | Should -BeFalse
        $notes.Text | Should -BeNullOrEmpty
        $notes.SlideIndex | Should -Be 0

        $presentationNotes = Get-OfficePowerPointNotes -Presentation $presentation -IncludeEmpty
        $presentationNotes.Count | Should -Be 1
        $presentationNotes[0].HasNotes | Should -BeFalse
        $slideSummary = Get-OfficePowerPointSlideSummary -Presentation $presentation
        $slideSummary.Count | Should -Be 1
        $slideSummary[0].HasNotes | Should -BeFalse
        $slideSummary[0].Title | Should -Be 'No notes yet'

        $shapeInfo = Get-OfficePowerPointShape -Presentation $presentation -Index 0 -Kind TextBox
        $shapeInfo.Count | Should -BeGreaterThan 0
        $shapeInfo[0].SlideIndex | Should -Be 0
        $shapeByIndex = Get-OfficePowerPointShape -Slide $slide -ShapeIndex $shapeInfo[0].ShapeIndex
        $shapeByIndex.ShapeIndex | Should -Be $shapeInfo[0].ShapeIndex
        $shapeByIndex.Kind | Should -Be 'TextBox'

        Save-OfficePowerPoint -Presentation $presentation

        $entries = Get-ZipEntriesLocal -Path $path
        ($entries | Where-Object { $_ -like 'ppt/notesSlides/*' }).Count | Should -Be 0
    }

    It 'persists layout placeholder edits across save and reopen' {
        $path = Join-Path $TestDrive 'PowerPointLayoutEdits.pptx'
        $presentation = New-OfficePowerPoint -FilePath $path
        $layoutPlaceholder = $null

        $layouts = Get-OfficePowerPointLayout -Presentation $presentation
        $layouts.Count | Should -BeGreaterThan 0

        $layoutType = $layouts | Where-Object { $_.Type } | Select-Object -First 1
        if ($layoutType) {
            $layoutMaster = $layoutType.MasterIndex
            $layoutIndex = $layoutType.LayoutIndex
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -LayoutType $layoutType.Type -Master $layoutType.MasterIndex
        } elseif ($layouts[0].Name) {
            $layoutMaster = $layouts[0].MasterIndex
            $layoutIndex = $layouts[0].LayoutIndex
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -LayoutName $layouts[0].Name -Master $layouts[0].MasterIndex
        } else {
            $layoutMaster = $layouts[0].MasterIndex
            $layoutIndex = $layouts[0].LayoutIndex
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout $layouts[0].LayoutIndex -Master $layouts[0].MasterIndex
        }

        $layoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder -Slide $slide
        $layoutPlaceholders.Count | Should -BeGreaterThan 0

        $layoutPlaceholder = $layoutPlaceholders | Where-Object { $_.PlaceholderType } | Select-Object -First 1
        if ($layoutPlaceholder) {
            $layoutPlaceholderType = $layoutPlaceholder.PlaceholderType.Value
            $boundsBox = Set-OfficePowerPointLayoutPlaceholderBounds -Presentation $presentation -Master $layoutMaster -Layout $layoutIndex -PlaceholderType $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex -Left 48 -Top 36 -Width 620 -Height 180 -PassThru
            $boundsBox.LeftPoints | Should -Be 48
            $boundsBox.TopPoints | Should -Be 36
            $boundsBox.WidthPoints | Should -Be 620
            $boundsBox.HeightPoints | Should -Be 180

            $marginsBox = Set-OfficePowerPointLayoutPlaceholderTextMargins -Presentation $presentation -Master $layoutMaster -Layout $layoutIndex -PlaceholderType $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex -Left 12 -Top 8 -Right 12 -Bottom 8 -PassThru
            $marginsBox.TextMarginLeftPoints | Should -Be 12
            $marginsBox.TextMarginTopPoints | Should -Be 8
            $marginsBox.TextMarginRightPoints | Should -Be 12
            $marginsBox.TextMarginBottomPoints | Should -Be 8

            $styleBox = Set-OfficePowerPointLayoutPlaceholderTextStyle -Presentation $presentation -Master $layoutMaster -Layout $layoutIndex -PlaceholderType $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex -Style Body -FontSize 18 -Bold $true -PassThru
            $styleBox.FontSize | Should -Be 18
            $styleBox.Bold | Should -BeTrue
        }

        Save-OfficePowerPoint -Presentation $presentation
        $presentation.Dispose()

        $reloaded = Get-OfficePowerPoint -FilePath $path
        try {
            $reloadedSlide = Get-OfficePowerPointSlide -Presentation $reloaded -Index 0
            $reloadedLayoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder -Slide $reloadedSlide
            if ($layoutPlaceholder) {
                $reloadedLayoutPlaceholders.Count | Should -BeGreaterThan 0
                $reloadedLayoutPlaceholder = Get-OfficePowerPointLayoutPlaceholder -Slide $reloadedSlide -PlaceholderType $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex
                $reloadedLayoutPlaceholder | Should -Not -BeNullOrEmpty
                $reloadedLayoutPlaceholder.Bounds | Should -Not -BeNullOrEmpty
                [math]::Abs($reloadedLayoutPlaceholder.Bounds.LeftPoints - 48) | Should -BeLessThan 0.1
                [math]::Abs($reloadedLayoutPlaceholder.Bounds.TopPoints - 36) | Should -BeLessThan 0.1
                [math]::Abs($reloadedLayoutPlaceholder.Bounds.WidthPoints - 620) | Should -BeLessThan 0.1
                [math]::Abs($reloadedLayoutPlaceholder.Bounds.HeightPoints - 180) | Should -BeLessThan 0.1
            }
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'removes slides and preserves the remaining title' {
        $path = Join-Path $TestDrive 'PowerPointRemoval.pptx'
        $presentation = New-OfficePowerPoint -FilePath $path

        $firstSlide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
        Set-OfficePowerPointSlideTitle -Slide $firstSlide -Title 'Remove me'

        $secondSlide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
        Set-OfficePowerPointSlideTitle -Slide $secondSlide -Title 'Keep me'
        Set-OfficePowerPointPlaceholderText -Slide $secondSlide -PlaceholderType Title -Text 'Keep me v2'

        Save-OfficePowerPoint -Presentation $presentation
        $presentation.Dispose()

        $reloaded = Get-OfficePowerPoint -FilePath $path
        try {
            Remove-OfficePowerPointSlide -Presentation $reloaded -Index 0 -Confirm:$false
            Save-OfficePowerPoint -Presentation $reloaded
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }

        $afterRemoval = Get-OfficePowerPoint -FilePath $path
        try {
            $afterRemoval.Slides.Count | Should -Be 1
            $remainingSlide = Get-OfficePowerPointSlide -Presentation $afterRemoval -Index 0
            $remainingPlaceholder = Get-OfficePowerPointPlaceholder -Slide $remainingSlide -PlaceholderType Title
            if (-not $remainingPlaceholder) {
                $remainingPlaceholder = Get-OfficePowerPointPlaceholder -Slide $remainingSlide -PlaceholderType CenteredTitle
            }
            $remainingPlaceholder.Text | Should -Be 'Keep me v2'
        } finally {
            if ($afterRemoval) {
                $afterRemoval.Dispose()
            }
        }
    }

    It 'copies a slide and preserves its content' {
        $path = Join-Path $TestDrive 'PowerPointCopy.pptx'
        $presentation = New-OfficePowerPoint -FilePath $path

        $slide1 = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
        Set-OfficePowerPointSlideTitle -Slide $slide1 -Title 'Quarterly Overview'
        Add-OfficePowerPointTextBox -Slide $slide1 -Text 'Revenue and margin summary' -X 80 -Y 150 -Width 320 -Height 60 | Out-Null
        Set-OfficePowerPointNotes -Slide $slide1 -Text 'Reuse this for the board deck.'

        $slide2 = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
        Set-OfficePowerPointSlideTitle -Slide $slide2 -Title 'Closing Slide'

        $copiedSlide = Copy-OfficePowerPointSlide -Presentation $presentation -Index 0 -InsertAt 1
        $copiedSlide | Should -Not -BeNullOrEmpty

        Save-OfficePowerPoint -Presentation $presentation
        $presentation.Dispose()

        $reloaded = Get-OfficePowerPoint -FilePath $path
        try {
            $reloaded.Slides.Count | Should -Be 3

            $copied = Get-OfficePowerPointSlide -Presentation $reloaded -Index 1
            $copiedTitle = Get-OfficePowerPointPlaceholder -Slide $copied -PlaceholderType Title
            if (-not $copiedTitle) {
                $copiedTitle = Get-OfficePowerPointPlaceholder -Slide $copied -PlaceholderType CenteredTitle
            }

            $copiedTitle.Text | Should -Be 'Quarterly Overview'
            (Get-OfficePowerPointNotes -Slide $copied).Text | Should -Be 'Reuse this for the board deck.'
            ($copied.Shapes.Count -gt 0) | Should -BeTrue

            $lastSlide = Get-OfficePowerPointSlide -Presentation $reloaded -Index 2
            $lastTitle = Get-OfficePowerPointPlaceholder -Slide $lastSlide -PlaceholderType Title
            if (-not $lastTitle) {
                $lastTitle = Get-OfficePowerPointPlaceholder -Slide $lastSlide -PlaceholderType CenteredTitle
            }

            $lastTitle.Text | Should -Be 'Closing Slide'
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'sets slide transitions and custom slide sizes' {
        $path = Join-Path $TestDrive 'PowerPointTransitionsAndSize.pptx'
        $presentation = New-OfficePowerPoint -FilePath $path

        $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
        Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Transition Demo' | Out-Null

        $updatedSlide = $slide | Set-OfficePowerPointSlideTransition -Transition Fade
        $updatedSlide.Transition | Should -Be ([OfficeIMO.PowerPoint.SlideTransition]::Fade)

        $slideSize = Set-OfficePowerPointSlideSize -Presentation $presentation -WidthCm 25.4 -HeightCm 14.0
        [math]::Round($slideSize.WidthCm, 1) | Should -Be 25.4
        [math]::Round($slideSize.HeightCm, 1) | Should -Be 14.0

        Save-OfficePowerPoint -Presentation $presentation
        $presentation.Dispose()

        $slideXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'ppt/slides/slide1.xml'
        $transitionNode = $slideXml.SelectSingleNode("/*[local-name()='sld']/*[local-name()='transition']")
        $transitionNode | Should -Not -BeNullOrEmpty
        $transitionNode.SelectSingleNode("*[local-name()='fade']") | Should -Not -BeNullOrEmpty

        $reloaded = Get-OfficePowerPoint -FilePath $path
        try {
            $reloadedSlide = Get-OfficePowerPointSlide -Presentation $reloaded -Index 0
            $reloadedSlide.Transition | Should -Be ([OfficeIMO.PowerPoint.SlideTransition]::Fade)
            [math]::Round($reloaded.SlideSize.WidthCm, 1) | Should -Be 25.4
            [math]::Round($reloaded.SlideSize.HeightCm, 1) | Should -Be 14.0
            $reloaded.SlideSize.IsLandscape | Should -BeTrue
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'applies preset slide sizes including portrait orientation' {
        $path = Join-Path $TestDrive 'PowerPointPresetSize.pptx'
        $presentation = New-OfficePowerPoint -FilePath $path
        Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 | Out-Null

        $presetSize = Set-OfficePowerPointSlideSize -Presentation $presentation -Preset Screen4x3 -Portrait
        $presetSize.IsPortrait | Should -BeTrue
        [math]::Round($presetSize.WidthInches, 1) | Should -Be 7.5
        [math]::Round($presetSize.HeightInches, 1) | Should -Be 10.0

        Save-OfficePowerPoint -Presentation $presentation
        $presentation.Dispose()

        $reloaded = Get-OfficePowerPoint -FilePath $path
        try {
            $reloaded.SlideSize.IsPortrait | Should -BeTrue
            [math]::Round($reloaded.SlideSize.WidthInches, 1) | Should -Be 7.5
            [math]::Round($reloaded.SlideSize.HeightInches, 1) | Should -Be 10.0
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'creates a presentation via the DSL context' {
        $path = Join-Path $TestDrive 'PowerPointDsl.pptx'

        New-OfficePowerPoint -Path $path {
            $layout = Get-OfficePowerPointLayout | Select-Object -First 1
            if ($layout) {
                Set-OfficePowerPointLayoutPlaceholderBounds -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title `
                    -Left 60 -Top 40 -Width 600 -Height 120 -CreateIfMissing
                Set-OfficePowerPointLayoutPlaceholderTextMargins -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title `
                    -Left 8 -Top 6 -Right 8 -Bottom 6 -CreateIfMissing
                Set-OfficePowerPointLayoutPlaceholderTextStyle -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title `
                    -Style Title -FontSize 32 -Bold $true -CreateIfMissing
            }

            PptSlide {
                PptTitle -Title 'DSL Slide'
                PptTextBox -Text 'Hello from DSL' -X 80 -Y 150 -Width 240 -Height 60
                $layoutPlaceholders = PptLayoutPlaceholders
                $layoutPlaceholders.Count | Should -BeGreaterThan 0
                $layoutPlaceholder = $layoutPlaceholders | Where-Object { $_.PlaceholderType } | Select-Object -First 1
                if ($layoutPlaceholder) {
                    $placeholder = Get-OfficePowerPointPlaceholder -PlaceholderType $layoutPlaceholder.PlaceholderType.Value -Index $layoutPlaceholder.PlaceholderIndex
                    if ($placeholder -and ((
                            $layoutPlaceholder.PlaceholderType.Value -eq [DocumentFormat.OpenXml.Presentation.PlaceholderValues]::Title
                        ) -or (
                            $layoutPlaceholder.PlaceholderType.Value -eq [DocumentFormat.OpenXml.Presentation.PlaceholderValues]::CenteredTitle
                        ))) {
                        $placeholder | Should -Not -BeNullOrEmpty
                        $placeholder.Text | Should -Be 'DSL Slide'
                    }
                }
            }
        }

        Test-Path $path | Should -BeTrue

        $presentation = Get-OfficePowerPoint -FilePath $path
        $slide = Get-OfficePowerPointSlide -Presentation $presentation -Index 0
        $summary = Get-OfficePowerPointSlideSummary -Slide $slide
        $summary.LayoutPlaceholderCount | Should -BeGreaterThan 0
        if ($summary.Title) {
            $summary.Title | Should -Be 'DSL Slide'
        }
        $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType Title
        if (-not $placeholder) {
            $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType CenteredTitle
        }
        if ($placeholder) {
            $placeholder.Text | Should -Be 'DSL Slide'
        }

        Save-OfficePowerPoint -Presentation $presentation
    }

    It 'supports sections, text replacement, and slide import' {
        $sourcePath = Join-Path $TestDrive 'PowerPointSourceDeck.pptx'
        $targetPath = Join-Path $TestDrive 'PowerPointTargetDeck.pptx'

        $source = New-OfficePowerPoint -FilePath $sourcePath
        $sourceSlide = Add-OfficePowerPointSlide -Presentation $source -Layout 1
        Set-OfficePowerPointSlideTitle -Slide $sourceSlide -Title 'FY24 Imported'
        Add-OfficePowerPointTextBox -Slide $sourceSlide -Text 'FY24 details from source' -X 80 -Y 150 -Width 320 -Height 60 | Out-Null
        Set-OfficePowerPointNotes -Slide $sourceSlide -Text 'FY24 source notes'
        Save-OfficePowerPoint -Presentation $source

        $target = New-OfficePowerPoint -FilePath $targetPath
        $introSlide = Add-OfficePowerPointSlide -Presentation $target -Layout 1
        Set-OfficePowerPointSlideTitle -Slide $introSlide -Title 'FY24 Overview'
        Add-OfficePowerPointTextBox -Slide $introSlide -Text 'FY24 summary for leadership' -X 80 -Y 150 -Width 320 -Height 60 | Out-Null
        Set-OfficePowerPointNotes -Slide $introSlide -Text 'FY24 note for intro'

        $resultsSlide = Add-OfficePowerPointSlide -Presentation $target -Layout 1
        Set-OfficePowerPointSlideTitle -Slide $resultsSlide -Title 'FY24 Results'
        Add-OfficePowerPointTextBox -Slide $resultsSlide -Text 'FY24 results body' -X 80 -Y 150 -Width 320 -Height 60 | Out-Null

        $introSection = Add-OfficePowerPointSection -Presentation $target -Name 'Intro' -StartSlideIndex 0
        $introSection.Name | Should -Be 'Intro'
        $resultsSection = Add-OfficePowerPointSection -Presentation $target -Name 'Results' -StartSlideIndex 1
        $resultsSection.Name | Should -Be 'Results'

        $renamedSection = Rename-OfficePowerPointSection -Presentation $target -Name 'Results' -NewName 'Deep Dive' -PassThru
        $renamedSection.Name | Should -Be 'Deep Dive'

        $sections = @(Get-OfficePowerPointSection -Presentation $target)
        $sections.Count | Should -Be 2
        ($sections | Where-Object Name -eq 'Intro').SlideIndices | Should -Contain 0
        ($sections | Where-Object Name -eq 'Deep Dive').SlideIndices | Should -Contain 1

        $replacements = Update-OfficePowerPointText -Presentation $target -OldValue 'FY24' -NewValue 'FY25' -IncludeNotes
        $replacements | Should -BeGreaterThan 0

        $importedSlide = Import-OfficePowerPointSlide -Presentation $target -SourcePath $sourcePath -SourceIndex 0 -InsertAt 1
        $importedSlide | Should -Not -BeNullOrEmpty

        Save-OfficePowerPoint -Presentation $target

        $reloaded = Get-OfficePowerPoint -FilePath $targetPath
        try {
            $reloaded.Slides.Count | Should -Be 3

            $slide0 = Get-OfficePowerPointSlide -Presentation $reloaded -Index 0
            $title0 = Get-OfficePowerPointPlaceholder -Slide $slide0 -PlaceholderType Title
            if (-not $title0) {
                $title0 = Get-OfficePowerPointPlaceholder -Slide $slide0 -PlaceholderType CenteredTitle
            }
            $title0.Text | Should -Be 'FY25 Overview'
            (Get-OfficePowerPointNotes -Slide $slide0).Text | Should -Be 'FY25 note for intro'

            $slide1 = Get-OfficePowerPointSlide -Presentation $reloaded -Index 1
            $title1 = Get-OfficePowerPointPlaceholder -Slide $slide1 -PlaceholderType Title
            if (-not $title1) {
                $title1 = Get-OfficePowerPointPlaceholder -Slide $slide1 -PlaceholderType CenteredTitle
            }
            $title1.Text | Should -Be 'FY24 Imported'
            (Get-OfficePowerPointNotes -Slide $slide1).Text | Should -Be 'FY24 source notes'

            $sectionInfo = @(Get-OfficePowerPointSection -Presentation $reloaded)
            $sectionInfo.Count | Should -Be 2
            ($sectionInfo | Where-Object Name -eq 'Intro').SlideIndices | Should -Contain 0
            ($sectionInfo | Where-Object Name -eq 'Intro').SlideIndices | Should -Contain 1
            ($sectionInfo | Where-Object Name -eq 'Deep Dive').SlideIndices | Should -Contain 2

            $aliasReplacements = Replace-OfficePowerPointText -Presentation $reloaded -OldValue 'FY24' -NewValue 'FY26'
            $aliasReplacements | Should -BeGreaterThan 0
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'supports theme inspection, theme updates, and slide layout switching' {
        $path = Join-Path $TestDrive 'PowerPointThemeAndLayout.pptx'
        $presentation = New-OfficePowerPoint -FilePath $path

        $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
        Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Theme Demo' | Out-Null

        $layouts = Get-OfficePowerPointLayout -Presentation $presentation
        $layouts.Count | Should -BeGreaterThan 1
        $currentLayoutIndex = $slide.LayoutIndex
        $alternativeLayout = $layouts | Where-Object { $_.LayoutIndex -ne $currentLayoutIndex } | Select-Object -First 1
        $alternativeLayout | Should -Not -BeNullOrEmpty

        $themeBefore = Get-OfficePowerPointTheme -Presentation $presentation
        $themeBefore.MasterIndex | Should -Be 0
        $themeBefore.Colors.Count | Should -BeGreaterThan 0

        Set-OfficePowerPointThemeColor -Presentation $presentation -Colors @{
            Accent1 = '#C00000'
            Accent2 = '00B0F0'
        } -AllMasters
        Set-OfficePowerPointThemeFonts -Presentation $presentation -MajorLatin 'Aptos' -MinorLatin 'Calibri' -AllMasters
        Set-OfficePowerPointThemeName -Presentation $presentation -Name 'Contoso Theme' -AllMasters

        if ($alternativeLayout.Type) {
            $slide | Set-OfficePowerPointSlideLayout -LayoutType $alternativeLayout.Type -Master $alternativeLayout.MasterIndex | Out-Null
        } elseif ($alternativeLayout.Name) {
            $slide | Set-OfficePowerPointSlideLayout -LayoutName $alternativeLayout.Name -Master $alternativeLayout.MasterIndex | Out-Null
        } else {
            $slide | Set-OfficePowerPointSlideLayout -Layout $alternativeLayout.LayoutIndex -Master $alternativeLayout.MasterIndex | Out-Null
        }

        $themeAfter = Get-OfficePowerPointTheme -Presentation $presentation
        $themeAfter.ThemeName | Should -Be 'Contoso Theme'
        $themeAfter.Colors[[OfficeIMO.PowerPoint.PowerPointThemeColor]::Accent1] | Should -Be 'C00000'
        $themeAfter.Colors[[OfficeIMO.PowerPoint.PowerPointThemeColor]::Accent2] | Should -Be '00B0F0'
        $themeAfter.MajorLatin | Should -Be 'Aptos'
        $themeAfter.MinorLatin | Should -Be 'Calibri'
        $slide.LayoutIndex | Should -Be $alternativeLayout.LayoutIndex

        Save-OfficePowerPoint -Presentation $presentation
        $presentation.Dispose()

        $reloaded = Get-OfficePowerPoint -FilePath $path
        try {
            $reloadedTheme = Get-OfficePowerPointTheme -Presentation $reloaded
            $reloadedTheme.ThemeName | Should -Be 'Contoso Theme'
            $reloadedTheme.Colors[[OfficeIMO.PowerPoint.PowerPointThemeColor]::Accent1] | Should -Be 'C00000'
            $reloadedTheme.Colors[[OfficeIMO.PowerPoint.PowerPointThemeColor]::Accent2] | Should -Be '00B0F0'
            $reloadedTheme.MajorLatin | Should -Be 'Aptos'
            $reloadedTheme.MinorLatin | Should -Be 'Calibri'

            $reloadedSlide = Get-OfficePowerPointSlide -Presentation $reloaded -Index 0
            $reloadedSlide.LayoutIndex | Should -Be $alternativeLayout.LayoutIndex
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }
}
