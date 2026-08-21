BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    function Get-OfficePrivateRegistryCount {
        param(
            [Parameter(Mandatory)]
            [string] $TypeName,
            [Parameter(Mandatory)]
            [string] $FieldName
        )

        $type = [AppDomain]::CurrentDomain.GetAssemblies() |
            ForEach-Object { $_.GetType($TypeName, $false) } |
            Where-Object { $null -ne $_ } |
            Select-Object -First 1
        $field = $type.GetField($FieldName, [Reflection.BindingFlags]'NonPublic,Static')
        $field.GetValue($null).Count
    }
}

Describe 'Office document PDF exports' {
    It 'exports saved native documents through one explicit command' {
        $wordPath = Join-Path $TestDrive 'new-word.docx'
        $wordPdf = Join-Path $TestDrive 'new-word.pdf'
        New-OfficeWord -Path $wordPath {
            WordSection {
                WordParagraph -Text 'New Word PDF sidecar'
            }
        }
        Export-OfficeDocumentPdf -InputPath $wordPath -Path $wordPdf

        $excelPath = Join-Path $TestDrive 'new-excel.xlsx'
        $excelPdf = Join-Path $TestDrive 'new-excel.pdf'
        New-OfficeExcel -Path $excelPath {
            ExcelSheet 'Data' {
                ExcelCell -Address 'A1' -Value 'New Excel PDF sidecar'
            }
        }
        Export-OfficeDocumentPdf -InputPath $excelPath -Path $excelPdf

        $markdownPath = Join-Path $TestDrive 'new-markdown.md'
        $markdownPdf = Join-Path $TestDrive 'new-markdown.pdf'
        New-OfficeMarkdown -Path $markdownPath {
            MarkdownHeading -Level 1 -Text 'New Markdown PDF sidecar'
        }
        Export-OfficeDocumentPdf -InputPath $markdownPath -Path $markdownPdf

        $powerPointPath = Join-Path $TestDrive 'new-powerpoint.pptx'
        $powerPointPdf = Join-Path $TestDrive 'new-powerpoint.pdf'
        New-OfficePowerPoint -Path $powerPointPath {
            PptSlide {
                PptTitle -Title 'New PowerPoint PDF sidecar'
            }
        }
        Export-OfficeDocumentPdf -InputPath $powerPointPath -Path $powerPointPdf

        foreach ($path in @($wordPath, $wordPdf, $excelPath, $excelPdf, $markdownPath, $markdownPdf, $powerPointPath, $powerPointPdf)) {
            Test-Path $path | Should -BeTrue
        }

        foreach ($path in @($wordPdf, $excelPdf, $markdownPdf, $powerPointPdf)) {
            (Get-OfficePdfPreflight -Path $path).CanRead | Should -BeTrue
        }
    }

    It 'exports an open Word document to PDF' {
        $docx = Join-Path $TestDrive 'word-report.docx'
        $pdf = Join-Path $TestDrive 'word-report.pdf'

        New-OfficeWord -Path $docx {
            WordSection {
                WordParagraph -Text 'Word PDF export smoke'
            }
        }

        $document = Get-OfficeWord -Path $docx
        try {
            $document | Export-OfficeDocumentPdf -Path $pdf
        } finally {
            Close-OfficeWord -Document $document
        }

        Test-Path $docx | Should -BeTrue
        Test-Path $pdf | Should -BeTrue
        (Get-OfficePdfPreflight -Path $pdf).CanRead | Should -BeTrue
        Get-OfficePdfText -Path $pdf | Should -Match 'Word PDF export smoke'
    }

    It 'releases encrypted path sources through their owning document services' {
        $password = 'pdf-export-secret'
        $wordPath = Join-Path $TestDrive 'encrypted-word.docx'
        $wordPdf = Join-Path $TestDrive 'encrypted-word.pdf'
        $excelPath = Join-Path $TestDrive 'encrypted-excel.xlsx'
        $excelPdf = Join-Path $TestDrive 'encrypted-excel.pdf'

        New-OfficeWord -Path $wordPath -Password $password {
            WordParagraph -Text 'Encrypted Word PDF export'
        }
        New-OfficeExcel -Path $excelPath -Password $password {
            ExcelSheet 'Data' {
                ExcelCell -Address A1 -Value 'Encrypted Excel PDF export'
            }
        }

        $wordBefore = Get-OfficePrivateRegistryCount -TypeName 'PSWriteOffice.Services.Word.WordDocumentService' -FieldName 'EncryptedSourcePaths'
        $excelBefore = Get-OfficePrivateRegistryCount -TypeName 'PSWriteOffice.Services.Excel.ExcelDocumentService' -FieldName 'EncryptedSourcePaths'

        Export-OfficeDocumentPdf -InputPath $wordPath -Path $wordPdf -Password $password
        Export-OfficeDocumentPdf -InputPath $excelPath -Path $excelPdf -Password $password

        Get-OfficePrivateRegistryCount -TypeName 'PSWriteOffice.Services.Word.WordDocumentService' -FieldName 'EncryptedSourcePaths' | Should -Be $wordBefore
        Get-OfficePrivateRegistryCount -TypeName 'PSWriteOffice.Services.Excel.ExcelDocumentService' -FieldName 'EncryptedSourcePaths' | Should -Be $excelBefore
        Test-Path -LiteralPath $wordPdf | Should -BeTrue
        Test-Path -LiteralPath $excelPdf | Should -BeTrue
    }

    It 'exports an open Excel workbook to PDF' {
        $xlsx = Join-Path $TestDrive 'excel-report.xlsx'
        $pdf = Join-Path $TestDrive 'excel-report.pdf'

        New-OfficeExcel -Path $xlsx {
            ExcelSheet 'Data' {
                ExcelCell -Address 'A1' -Value 'Excel PDF export smoke'
                ExcelCell -Address 'B1' -Value 42
                ExcelAutoFit
            }
        }

        $workbook = Get-OfficeExcel -Path $xlsx
        try {
            $workbook | Export-OfficeDocumentPdf -Path $pdf
        } finally {
            Close-OfficeExcel -Document $workbook
        }

        Test-Path $xlsx | Should -BeTrue
        Test-Path $pdf | Should -BeTrue
        (Get-OfficePdfPreflight -Path $pdf).CanRead | Should -BeTrue
        Get-OfficePdfText -Path $pdf | Should -Match 'Excel PDF export smoke'
    }

    It 'exports a Markdown document to PDF' {
        $md = Join-Path $TestDrive 'markdown-report.md'
        $pdf = Join-Path $TestDrive 'markdown-pdf\markdown-report.pdf'
        $document = Get-OfficeMarkdown -Text "# Markdown PDF export smoke`n`nGenerated by PSWriteOffice."

        $document | Save-OfficeMarkdown -Path $md
        $document | Export-OfficeDocumentPdf -Path $pdf

        Test-Path $md | Should -BeTrue
        Test-Path $pdf | Should -BeTrue
        (Get-OfficePdfPreflight -Path $pdf).CanRead | Should -BeTrue
        Get-OfficePdfText -Path $pdf | Should -Match 'Markdown PDF export smoke'
    }

    It 'exports an open PowerPoint presentation to PDF' {
        $pptx = Join-Path $TestDrive 'deck-report.pptx'
        $pdf = Join-Path $TestDrive 'deck-report.pdf'

        New-OfficePowerPoint -Path $pptx {
            PptSlide {
                PptTitle -Title 'PowerPoint PDF export smoke'
            }
        }

        $presentation = Get-OfficePowerPoint -Path $pptx
        try {
            $presentation | Export-OfficeDocumentPdf -Path $pdf
        } finally {
            $presentation | Close-OfficePowerPoint
        }

        Test-Path $pptx | Should -BeTrue
        Test-Path $pdf | Should -BeTrue
        $info = Get-OfficePdfInfo -Path $pdf
        $info.PageCount | Should -BeGreaterOrEqual 1
        (Get-OfficePdfPreflight -Path $pdf).CanRead | Should -BeTrue
    }
}
