Describe 'Get-OfficeWord cmdlet' {
    It 'throws when file does not exist' {
        { Get-OfficeWord -FilePath (Join-Path $TestDrive 'missing.docx') -ErrorAction Stop } | Should -Throw
    }

    It 'loads existing Word document' {
        $path = Join-Path $TestDrive 'test.docx'
        $newDoc = New-OfficeWord -FilePath $path
        Save-OfficeWord -Document $newDoc
        $doc = Get-OfficeWord -FilePath $path
        $doc | Should -Not -BeNullOrEmpty
        Close-OfficeWord -Document $doc
    }
}
