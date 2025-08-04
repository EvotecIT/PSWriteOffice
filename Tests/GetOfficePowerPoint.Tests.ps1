Describe 'Get-OfficePowerPoint cmdlet' {
    It 'throws when file does not exist' {
        { Get-OfficePowerPoint -FilePath (Join-Path $TestDrive 'missing.pptx') -ErrorAction Stop } | Should -Throw
    }

    It 'loads existing presentation' {
        $path = Join-Path $TestDrive 'test.pptx'
        $pres = New-OfficePowerPoint -FilePath $path
        Save-OfficePowerPoint -Presentation $pres
        $loaded = Get-OfficePowerPoint -FilePath $path
        $loaded | Should -Not -BeNullOrEmpty
    }
}
