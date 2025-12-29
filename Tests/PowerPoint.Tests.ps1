BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force
}

Describe 'PowerPoint cmdlets' {
    It 'creates a presentation with shapes and tables' {
        $path = Join-Path $TestDrive 'PowerPointBasics.pptx'
        $presentation = New-OfficePowerPoint -FilePath $path

        $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
        $shape = Add-OfficePowerPointShape -Slide $slide -ShapeType Rectangle -X 40 -Y 40 -Width 200 -Height 80 -FillColor '#DDEEFF' -OutlineColor '#1F4E79' -OutlineWidth 1

        $rows = @(
            [PSCustomObject]@{ Item = 'Alpha'; Qty = 10 }
            [PSCustomObject]@{ Item = 'Beta'; Qty = 20 }
        )
        $table = Add-OfficePowerPointTable -Slide $slide -Data $rows -X 40 -Y 140 -Width 360 -Height 200

        Save-OfficePowerPoint -Presentation $presentation

        Test-Path $path | Should -BeTrue
        $slide.Shapes.Count | Should -BeGreaterThan 0
        @($slide.Tables).Count | Should -BeGreaterThan 0
        $table.Rows | Should -BeGreaterThan 0
        $shape | Should -Not -BeNullOrEmpty
    }
}
