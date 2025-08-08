Describe 'Remove-OfficePowerPointSlide cmdlet' {
    It 'removes slide at index' {
        $path = Join-Path $TestDrive 'remove.pptx'
        $pres = New-OfficePowerPoint -FilePath $path
        Add-OfficePowerPointSlide -Presentation $pres -Layout 1 | Out-Null
        Add-OfficePowerPointSlide -Presentation $pres -Layout 1 | Out-Null
        Remove-OfficePowerPointSlide -Presentation $pres -Index 0
        $pres.Slides.Count | Should -Be 1
    }

    It 'does not remove slide when using -WhatIf' {
        $path = Join-Path $TestDrive 'removeWhatIf.pptx'
        $pres = New-OfficePowerPoint -FilePath $path
        Add-OfficePowerPointSlide -Presentation $pres -Layout 1 | Out-Null
        Add-OfficePowerPointSlide -Presentation $pres -Layout 1 | Out-Null
        Remove-OfficePowerPointSlide -Presentation $pres -Index 0 -WhatIf
        $pres.Slides.Count | Should -Be 2
    }
}
