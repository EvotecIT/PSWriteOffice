Describe 'PowerPoint cmdlets' {
    It 'creates new presentation' {
        $path = Join-Path $TestDrive 'test.pptx'
        $pres = New-OfficePowerPoint -FilePath $path
        $pres | Should -Not -BeNullOrEmpty
    }

    It 'saves presentation to path' {
        $path = Join-Path $TestDrive 'save.pptx'
        $pres = New-OfficePowerPoint -FilePath $path
        Save-OfficePowerPoint -Presentation $pres
        Test-Path $path | Should -BeTrue
    }

    It 'supports WhatIf when saving presentation' {
        $path = Join-Path $TestDrive 'whatif.pptx'
        $pres = New-OfficePowerPoint -FilePath $path
        Save-OfficePowerPoint -Presentation $pres -WhatIf
        Test-Path $path | Should -BeFalse
    }

    It 'adds slide to presentation' {
        $path = Join-Path $TestDrive 'addslide.pptx'
        $pres = New-OfficePowerPoint -FilePath $path
        Add-OfficePowerPointSlide -Presentation $pres -Layout 1
        # ShapeCrawler always maintains at least 1 slide, so we expect 2
        $pres.Slides.Count | Should -Be 2
    }

    It 'merges presentations' {
        $targetPath = Join-Path $TestDrive 'target.pptx'
        $sourcePath = Join-Path $TestDrive 'source.pptx'
        $target = New-OfficePowerPoint -FilePath $targetPath
        $source = New-OfficePowerPoint -FilePath $sourcePath
        Add-OfficePowerPointSlide -Presentation $source -Layout 1
        Save-OfficePowerPoint -Presentation $source
        Merge-OfficePowerPoint -Presentation $target -FilePath $sourcePath
        # Target starts with 1, source has 2 (1 default + 1 added), so merged = 3
        $target.Slides.Count | Should -Be 3
    }

    It 'gets slides by index' {
        $path = Join-Path $TestDrive 'getslide.pptx'
        $pres = New-OfficePowerPoint -FilePath $path
        Add-OfficePowerPointSlide -Presentation $pres -Layout 1 | Out-Null
        $slide = Get-OfficePowerPointSlide -Presentation $pres -Index 0
        $slide | Should -Not -BeNullOrEmpty
    }

    It 'adds textbox to slide' {
        $path = Join-Path $TestDrive 'textbox.pptx'
        $pres = New-OfficePowerPoint -FilePath $path
        $slide = Add-OfficePowerPointSlide -Presentation $pres -Layout 1
        Add-OfficePowerPointTextBox -Slide $slide -Text 'Hello'
        $slide.Shapes[$slide.Shapes.Count - 1].TextBox.Text | Should -Be 'Hello'
    }

    It 'sets slide title' {
        $path = Join-Path $TestDrive 'title.pptx'
        $pres = New-OfficePowerPoint -FilePath $path
        $slide = Add-OfficePowerPointSlide -Presentation $pres -Layout 1
        Set-OfficePowerPointSlideTitle -Slide $slide -Title 'My Title'
        $slide.Shapes.Shape('Title 1').TextBox.Text | Should -Be 'My Title'
    }
}
