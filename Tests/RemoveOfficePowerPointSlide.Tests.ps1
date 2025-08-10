Describe 'Remove-OfficePowerPointSlide cmdlet' {
    It 'removes slide at index' {
        $path = Join-Path $TestDrive 'remove.pptx'
        $pres = New-OfficePowerPoint -FilePath $path
        Add-OfficePowerPointSlide -Presentation $pres -Layout 1 | Out-Null
        Add-OfficePowerPointSlide -Presentation $pres -Layout 1 | Out-Null
        Remove-OfficePowerPointSlide -Presentation $pres -Index 0
        # Started with 1 default, added 2 (total 3), removed 1, so expect 2
        $pres.Slides.Count | Should -Be 2
    }

    It 'does not remove slide when using -WhatIf' {
        $path = Join-Path $TestDrive 'removeWhatIf.pptx'
        $pres = New-OfficePowerPoint -FilePath $path
        Add-OfficePowerPointSlide -Presentation $pres -Layout 1 | Out-Null
        Add-OfficePowerPointSlide -Presentation $pres -Layout 1 | Out-Null
        Remove-OfficePowerPointSlide -Presentation $pres -Index 0 -WhatIf
        # Started with 1 default, added 2 (total 3), no removal due to WhatIf
        $pres.Slides.Count | Should -Be 3
    }
}
