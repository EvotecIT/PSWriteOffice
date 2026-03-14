BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Force -Global
}

Describe 'PowerPoint cmdlets' {
    It 'creates a presentation with shapes, tables, and media' {
        $path = Join-Path $TestDrive 'PowerPointBasics.pptx'
        $presentation = New-OfficePowerPoint -FilePath $path
        $officeimoRoot = Join-Path $PSScriptRoot '..\..\OfficeIMO'
        $imagePath = Join-Path (Join-Path $officeimoRoot 'Assets') 'OfficeIMO.png'
        if (-not (Test-Path $imagePath)) {
            throw "OfficeIMO image asset not found at $imagePath"
        }

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
        $slide.Notes.Text | Should -Be 'Keep this under five minutes.'
        $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType Title
        $placeholder.Text | Should -Be 'Status Update'
        $placeholderUpdate = Set-OfficePowerPointPlaceholderText -Slide $slide -PlaceholderType Title -Text 'Status Update v2' -PassThru
        $placeholderUpdate.Text | Should -Be 'Status Update v2'
        $layoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder -Slide $slide
        $layoutPlaceholders.Count | Should -BeGreaterThan 0

        $layoutPlaceholdersForLayout = Get-OfficePowerPointLayoutPlaceholder -Slide $slide2
        $layoutPlaceholdersForLayout.Count | Should -BeGreaterThan 0
        $layoutPlaceholder = $layoutPlaceholdersForLayout | Where-Object { $_.PlaceholderType } | Select-Object -First 1
        if ($layoutPlaceholder) {
            $layoutPlaceholderType = $layoutPlaceholder.PlaceholderType.Value
            $boundsBox = Set-OfficePowerPointLayoutPlaceholderBounds -Presentation $presentation -Master $layoutMaster -Layout $layoutIndex -PlaceholderType $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex -Left 48 -Top 36 -Width 620 -Height 180 -PassThru
            $boundsBox | Should -Not -BeNullOrEmpty
            $marginsBox = Set-OfficePowerPointLayoutPlaceholderTextMargins -Presentation $presentation -Master $layoutMaster -Layout $layoutIndex -PlaceholderType $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex -Left 12 -Top 8 -Right 12 -Bottom 8 -PassThru
            $marginsBox | Should -Not -BeNullOrEmpty
            $styleBox = Set-OfficePowerPointLayoutPlaceholderTextStyle -Presentation $presentation -Master $layoutMaster -Layout $layoutIndex -PlaceholderType $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex -Style Body -FontSize 18 -Bold $true -PassThru
            $styleBox | Should -Not -BeNullOrEmpty
        }
        $slide2 | Should -Not -BeNullOrEmpty

        Save-OfficePowerPoint -Presentation $presentation

        Test-Path $path | Should -BeTrue
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
                    $placeholder | Should -Not -BeNullOrEmpty
                    if ($layoutPlaceholder.PlaceholderType.Value -eq [DocumentFormat.OpenXml.Presentation.PlaceholderValues]::Title -or
                        $layoutPlaceholder.PlaceholderType.Value -eq [DocumentFormat.OpenXml.Presentation.PlaceholderValues]::CenteredTitle) {
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
