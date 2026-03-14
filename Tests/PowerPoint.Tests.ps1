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
        $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType Title
        if (-not $placeholder) {
            $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType CenteredTitle
        }
        if ($placeholder) {
            $placeholder.Text | Should -Be 'DSL Slide'
        }

        Save-OfficePowerPoint -Presentation $presentation
    }
}
