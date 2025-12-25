BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force
}

Describe 'CSV cmdlets' {
    It 'converts objects to CSV and reads them back' {
        $rows = @(
            [pscustomobject]@{ Region = 'NA'; Revenue = 100 }
            [pscustomobject]@{ Region = 'EMEA'; Revenue = 200 }
        )

        $csvText = $rows | ConvertTo-OfficeCsv
        $csvText | Should -Match 'Region'
        $csvText | Should -Match 'Revenue'

        $path = Join-Path $TestDrive 'data.csv'
        $rows | ConvertTo-OfficeCsv -OutputPath $path | Out-Null

        Test-Path $path | Should -BeTrue

        $data = Get-OfficeCsvData -Path $path
        $data.Count | Should -Be 2
        $data[0].Region | Should -Be 'NA'
    }
}
