BeforeAll {
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $moduleManifest -Global -ErrorAction Stop
}

Describe 'Object-style document composition' {
    It 'composes and saves a Word document without a DSL scriptblock' {
        $path = Join-Path $TestDrive 'ObjectFlow.docx'
        $rows = @(
            [pscustomobject]@{ Service = 'Identity'; Status = 'Ready' }
            [pscustomobject]@{ Service = 'Certificates'; Status = 'Watch' }
        )

        $document = New-OfficeWord -Path $path -NoSave
        try {
            $heading = $document | Add-OfficeWordParagraph -Text 'Service status' -Style Heading1 -PassThru
            $heading | Add-OfficeWordText -Text ' — reviewed' -Italic

            $details = $document | Add-OfficeWordParagraph -PassThru
            $details | Add-OfficeWordText -Run @{
                Text  = 'Owner: ', 'Platform', '  Status: ', 'Ready'
                Bold  = $true, $false, $true, $true
                Color = $null, $null, $null, 'SeaGreen'
            }

            Add-OfficeWordTable -Document $document -InputObject $rows -Style GridTable4Accent1 -Layout AutoFitToWindow
            $document | Save-OfficeWord
        } finally {
            $document | Close-OfficeWord
        }

        $path | Should -Exist
        $readBack = Get-OfficeWord -Path $path -ReadOnly
        try {
            ($readBack.Paragraphs.Text -join '') | Should -Match 'Service status — reviewed'
            $readBack.Tables.Count | Should -Be 1
        } finally {
            $readBack | Close-OfficeWord
        }
    }

    It 'adds a Word section through the document pipeline' {
        $path = Join-Path $TestDrive 'ObjectSections.docx'
        $document = New-OfficeWord -Path $path -NoSave
        try {
            $section = $document | Add-OfficeWordSection -BreakType NextPage -PassThru
            $section | Add-OfficeWordParagraph -Text 'Second section' -PassThru | Should -Not -BeNullOrEmpty
            $document | Save-OfficeWord
        } finally {
            $document | Close-OfficeWord
        }

        $readBack = Get-OfficeWord -Path $path -ReadOnly
        try {
            $readBack.Sections.Count | Should -BeGreaterThan 1
            ($readBack.Paragraphs.Text -join "`n") | Should -Match 'Second section'
        } finally {
            $readBack | Close-OfficeWord
        }
    }

    It 'composes and saves an Excel workbook without a DSL scriptblock' {
        $path = Join-Path $TestDrive 'ObjectFlow.xlsx'
        $rows = @(
            [pscustomobject]@{ Project = 'Atlas'; Owner = 'Operations'; Progress = 0.80 }
            [pscustomobject]@{ Project = 'Beacon'; Owner = 'Security'; Progress = 0.55 }
        )

        $workbook = New-OfficeExcel -Path $path -NoSave
        try {
            $sheet = $workbook | Add-OfficeExcelSheet -Name 'Projects' -PassThru
            $sheet | Set-OfficeExcelCell -Address A1 -Value 'Portfolio' -BackgroundColor '#D9EAF7'
            Add-OfficeExcelTable -Worksheet $sheet -InputObject $rows -StartRow 3 -TableName 'Projects' -AutoFit
            Set-OfficeExcelCell -Document $workbook -Sheet 'Projects' -Address C4 -NumberFormat '0%'
            $workbook | Save-OfficeExcel
        } finally {
            $workbook | Close-OfficeExcel
        }

        $path | Should -Exist
        $data = @(Import-OfficeExcel -Path $path -WorksheetName 'Projects' -Range 'A3:C5')
        $data.Count | Should -Be 2
        $data[0].Project | Should -Be 'Atlas'
    }
}
