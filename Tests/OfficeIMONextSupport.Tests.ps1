BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Expanded OfficeIMO support' {
    It 'exports every new command family' {
        $commands = @(
            'Get-OfficeDocumentDetection', 'Get-OfficeDocumentStructured', 'Get-OfficeDocumentHierarchy',
            'Get-OfficeDocumentBatch', 'New-OfficeDocumentReader',
            'Export-OfficeWordImage', 'Export-OfficeExcelImage', 'Export-OfficePowerPointImage',
            'Export-OfficeHtmlImage', 'Export-OfficePdfImage',
            'Compare-OfficeWordDocument', 'Get-OfficeWordReview', 'Resolve-OfficeWordRevision',
            'Get-OfficePowerPointInspection', 'ConvertTo-OfficePdfSanitized',
            'Export-OfficePdfXfdf', 'Import-OfficePdfXfdf', 'Compare-OfficePdfVisual',
            'Get-OfficePdfInteractionMap', 'Export-OfficePdfLayoutOverlay', 'Invoke-OfficePdfOcrMerge',
            'Test-OfficePdfRewrite', 'New-OfficeOpenDocument', 'Get-OfficeOpenDocument',
            'Save-OfficeOpenDocument', 'ConvertTo-OfficeOpenDocument', 'ConvertFrom-OfficeOpenDocument',
            'Get-OfficeEmail', 'Save-OfficeEmail', 'Get-OfficeEmailMailbox', 'Save-OfficeEmailMailbox',
            'Get-OfficeAsciiDoc', 'Save-OfficeAsciiDoc', 'ConvertTo-OfficeAsciiDocMarkdown',
            'ConvertFrom-OfficeAsciiDocMarkdown', 'Get-OfficeLatex', 'Save-OfficeLatex',
            'ConvertTo-OfficeLatexMarkdown', 'ConvertFrom-OfficeLatexMarkdown',
            'Export-OfficeWordGoogleDocument', 'Export-OfficeExcelGoogleSpreadsheet'
        )

        foreach ($command in $commands) {
            Get-Command $command -ErrorAction Stop | Should -Not -BeNullOrEmpty
        }
        (Get-Command New-OfficeDocumentReader).Parameters.Keys | Should -Contain 'TesseractOptions'
        (Get-Command New-OfficeDocumentReader).Parameters.Keys | Should -Contain 'ProcessOcrOptions'
        (Get-Command Invoke-OfficePdfOcrMerge).Parameters.Keys | Should -Contain 'Provider'
    }

    It 'detects, structures, chunks, and batch-reads additional text formats' {
        $asciiDocPath = Join-Path $TestDrive 'guide.adoc'
        $latexPath = Join-Path $TestDrive 'paper.tex'
        Set-Content -LiteralPath $asciiDocPath -Encoding UTF8 -Value "= Guide`n`n== Reader`n`nStructured AsciiDoc content."
        Set-Content -LiteralPath $latexPath -Encoding UTF8 -Value '\documentclass{article}\begin{document}\section{Reader}Structured LaTeX content.\end{document}'

        $detection = Get-OfficeDocumentDetection -Path $asciiDocPath
        $detection.GetType().FullName | Should -Be 'OfficeIMO.Reader.ReaderDetectionResult'

        $structured = Get-OfficeDocumentStructured -Path $asciiDocPath
        $structured.GetType().FullName | Should -Be 'OfficeIMO.Reader.OfficeDocumentStructuredExtractionResult'
        $structured.Records.Count | Should -BeGreaterThan 0

        $hierarchy = Get-OfficeDocumentHierarchy -Path $asciiDocPath
        $hierarchy.GetType().FullName | Should -Be 'OfficeIMO.Reader.ReaderChunkHierarchyResult'
        $hierarchy.Chunks.Count | Should -BeGreaterThan 0

        $batch = @(@($asciiDocPath, $latexPath) | Get-OfficeDocumentBatch)
        $batch | Should -HaveCount 2
        $batch[0].GetType().FullName | Should -Be 'OfficeIMO.Reader.OfficeDocumentReadResult'

        $reader = New-OfficeDocumentReader -MaxConcurrentReads 2
        $reader.GetType().FullName | Should -Be 'OfficeIMO.Reader.OfficeDocumentReader'
    }

    It 'round-trips AsciiDoc and LaTeX through loss-aware Markdown results' {
        $markdown = Get-OfficeMarkdown -Text "# Portable report`n`nA compact conversion contract."
        $asciiDocPath = Join-Path $TestDrive 'portable.adoc'
        $latexPath = Join-Path $TestDrive 'portable.tex'

        $asciiDocResult = $markdown | ConvertFrom-OfficeAsciiDocMarkdown -OutputPath $asciiDocPath -FailOnLoss
        $asciiDocResult.GetType().FullName | Should -Be 'OfficeIMO.AsciiDoc.Markdown.MarkdownToAsciiDocResult'
        $asciiDocResult.HasLoss | Should -BeFalse
        Test-Path -LiteralPath $asciiDocPath | Should -BeTrue
        (Get-OfficeAsciiDoc -Path $asciiDocPath -AsResult).IsLossless | Should -BeTrue
        $asciiDocMarkdown = ConvertTo-OfficeAsciiDocMarkdown -Path $asciiDocPath -FailOnLoss
        $asciiDocMarkdown.Value.ToMarkdown() | Should -Match 'Portable report'

        $latexResult = $markdown | ConvertFrom-OfficeLatexMarkdown -OutputPath $latexPath -FailOnLoss
        $latexResult.GetType().FullName | Should -Be 'OfficeIMO.Latex.Markdown.MarkdownToLatexResult'
        $latexResult.HasLoss | Should -BeFalse
        Test-Path -LiteralPath $latexPath | Should -BeTrue
        (Get-OfficeLatex -Path $latexPath -AsResult).IsLossless | Should -BeTrue
        $latexMarkdown = ConvertTo-OfficeLatexMarkdown -Path $latexPath -FailOnLoss
        $latexMarkdown.Value.ToMarkdown() | Should -Match 'Portable report'
    }

    It 'creates and reloads native ODT, ODS, and ODP packages' {
        $cases = @(
            @{ Kind = 'Text'; Extension = 'odt'; Type = 'OfficeIMO.OpenDocument.OdtDocument' },
            @{ Kind = 'Spreadsheet'; Extension = 'ods'; Type = 'OfficeIMO.OpenDocument.OdsDocument' },
            @{ Kind = 'Presentation'; Extension = 'odp'; Type = 'OfficeIMO.OpenDocument.OdpPresentation' }
        )

        foreach ($case in $cases) {
            $path = Join-Path $TestDrive "native.$($case.Extension)"
            $document = New-OfficeOpenDocument -Kind $case.Kind
            $save = $document | Save-OfficeOpenDocument -Path $path -FailOnLoss
            $save.HasLoss | Should -BeFalse
            Test-Path -LiteralPath $path | Should -BeTrue
            (Get-OfficeOpenDocument -Path $path).GetType().FullName | Should -Be $case.Type
        }
    }

    It 'round-trips EML, MSG, TNEF, and mbox artifacts' {
        $emailDocumentType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Email' -TypeName 'OfficeIMO.Email.EmailDocument' -CommandName 'Save-OfficeEmail'
        $emailAddressType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Email' -TypeName 'OfficeIMO.Email.EmailAddress' -CommandName 'Save-OfficeEmail'
        $emailMailboxType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Email' -TypeName 'OfficeIMO.Email.EmailMailbox' -CommandName 'Save-OfficeEmailMailbox'
        $emailMailboxEntryType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Email' -TypeName 'OfficeIMO.Email.EmailMailboxEntry' -CommandName 'Save-OfficeEmailMailbox'
        $message = [Activator]::CreateInstance($emailDocumentType)
        $message.Subject = 'Native email contract'
        $message.From = [Activator]::CreateInstance($emailAddressType, @('sender@example.test', 'Sender', $null))
        $message.Body.Text = 'OfficeIMO email body'

        foreach ($extension in 'eml', 'msg', 'dat') {
            $path = Join-Path $TestDrive "message.$extension"
            $format = if ($extension -eq 'dat') { 'Tnef' } else { $null }
            $result = if ($format) {
                $message | Save-OfficeEmail -Path $path -Format $format
            } else {
                $message | Save-OfficeEmail -Path $path
            }
            $result.GetType().FullName | Should -Be 'OfficeIMO.Email.EmailWriteResult'
            Test-Path -LiteralPath $path | Should -BeTrue
            (Get-OfficeEmail -Path $path).Subject | Should -Be 'Native email contract'
        }

        $mailbox = [Activator]::CreateInstance($emailMailboxType)
        $mailbox.Messages.Add([Activator]::CreateInstance($emailMailboxEntryType, @($message)))
        $mailboxPath = Join-Path $TestDrive 'mailbox.mbox'
        $mailboxResult = $mailbox | Save-OfficeEmailMailbox -Path $mailboxPath
        $mailboxResult.GetType().FullName | Should -Be 'OfficeIMO.Email.EmailWriteResult'
        (Get-OfficeEmailMailbox -Path $mailboxPath).Messages.Count | Should -Be 1
    }

    It 'returns native image export results for Office, HTML, and PDF inputs' {
        $wordPath = Join-Path $TestDrive 'images.docx'
        $excelPath = Join-Path $TestDrive 'images.xlsx'
        $powerPointPath = Join-Path $TestDrive 'images.pptx'
        $pdfPath = Join-Path $TestDrive 'images.pdf'
        New-OfficeWord -Path $wordPath { WordSection { WordParagraph -Text 'Word image' } } | Out-Null
        New-OfficeExcel -Path $excelPath { ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Excel image' } } | Out-Null
        New-OfficePowerPoint -Path $powerPointPath { PptSlide { PptTitle -Title 'PowerPoint image' } } | Out-Null
        New-OfficePdf -Path $pdfPath { PdfParagraph 'PDF image' } | Out-Null

        $results = @(
            Export-OfficeWordImage -Path $wordPath -OutputPath (Join-Path $TestDrive 'word.svg') -Format Svg
            Export-OfficeExcelImage -Path $excelPath -OutputPath (Join-Path $TestDrive 'excel-images') -Format Svg
            Export-OfficePowerPointImage -Path $powerPointPath -OutputPath (Join-Path $TestDrive 'ppt-images') -Format Svg
            Export-OfficeHtmlImage -Html '<h1>HTML image</h1>' -OutputPath (Join-Path $TestDrive 'html.svg') -Format Svg
            Export-OfficePdfImage -Path $pdfPath -OutputPath (Join-Path $TestDrive 'pdf-images') -Format Svg
        )

        $results.Count | Should -BeGreaterOrEqual 5
        foreach ($result in $results) {
            $result.GetType().FullName | Should -Be 'OfficeIMO.Drawing.OfficeImageExportResult'
            $result.Bytes.Length | Should -BeGreaterThan 0
        }
    }

    It 'returns Word review/comparison and PowerPoint inspection reports' {
        $before = Join-Path $TestDrive 'before.docx'
        $after = Join-Path $TestDrive 'after.docx'
        $redline = Join-Path $TestDrive 'redline.docx'
        $deck = Join-Path $TestDrive 'inspection.pptx'
        New-OfficeWord -Path $before { WordSection { WordParagraph -Text 'Before text' } } | Out-Null
        New-OfficeWord -Path $after { WordSection { WordParagraph -Text 'After text' } } | Out-Null
        New-OfficePowerPoint -Path $deck { PptSlide { PptTitle -Title 'Inspection' } } | Out-Null

        $comparison = Compare-OfficeWordDocument -ReferencePath $before -DifferencePath $after -RedlinePath $redline
        $comparison.GetType().FullName | Should -Be 'OfficeIMO.Word.WordComparisonResult'
        Test-Path -LiteralPath $redline | Should -BeTrue
        (Get-OfficeWordReview -Path $redline).GetType().FullName | Should -Be 'OfficeIMO.Word.WordReviewReport'
        (Get-OfficePowerPointInspection -Path $deck).GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointInspectionReport'
    }

    It 'runs PDF sanitization, XFDF, interaction, overlay, comparison, and rewrite proof' {
        $pdfPath = Join-Path $TestDrive 'advanced.pdf'
        $safePath = Join-Path $TestDrive 'advanced-safe.pdf'
        $overlayPath = Join-Path $TestDrive 'advanced-overlay.svg'
        New-OfficePdf -Path $pdfPath { PdfParagraph 'Advanced PDF proof' } | Out-Null

        $sanitized = ConvertTo-OfficePdfSanitized -Path $pdfPath -OutputPath $safePath
        $sanitized.GetType().FullName | Should -Be 'OfficeIMO.Pdf.PdfSanitizationResult'
        Test-Path -LiteralPath $safePath | Should -BeTrue
        (Export-OfficePdfXfdf -Path $pdfPath) | Should -Match '<xfdf'
        (Get-OfficePdfInteractionMap -Path $pdfPath).GetType().FullName | Should -Be 'OfficeIMO.Pdf.PdfPageInteractionMap'
        (Export-OfficePdfLayoutOverlay -Path $pdfPath -OutputPath $overlayPath).GetType().FullName |
            Should -Be 'OfficeIMO.Drawing.OfficeImageExportResult'
        Test-Path -LiteralPath $overlayPath | Should -BeTrue
        (Compare-OfficePdfVisual -ReferencePath $pdfPath -DifferencePath $pdfPath).GetType().FullName |
            Should -Be 'OfficeIMO.Pdf.PdfVisualComparisonReport'
        (Test-OfficePdfRewrite -ReferencePath $pdfPath -DifferencePath $safePath).GetType().FullName |
            Should -Be 'OfficeIMO.Pdf.PdfRewritePreservationReport'
    }

    It 'builds offline Google Docs and Sheets plans and request batches' {
        $wordPath = Join-Path $TestDrive 'google.docx'
        $excelPath = Join-Path $TestDrive 'google.xlsx'
        New-OfficeWord -Path $wordPath { WordSection { WordParagraph -Text 'Google plan' } } | Out-Null
        New-OfficeExcel -Path $excelPath { ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Google plan' } } | Out-Null

        (Export-OfficeWordGoogleDocument -Path $wordPath -PlanOnly).GetType().FullName |
            Should -Be 'OfficeIMO.Word.GoogleDocs.GoogleDocsTranslationPlan'
        (Export-OfficeWordGoogleDocument -Path $wordPath -AsBatch).GetType().FullName |
            Should -Be 'OfficeIMO.Word.GoogleDocs.GoogleDocsBatch'
        (Export-OfficeExcelGoogleSpreadsheet -Path $excelPath -PlanOnly).GetType().FullName |
            Should -Be 'OfficeIMO.Excel.GoogleSheets.GoogleSheetsTranslationPlan'
        (Export-OfficeExcelGoogleSpreadsheet -Path $excelPath -AsBatch).GetType().FullName |
            Should -Be 'OfficeIMO.Excel.GoogleSheets.GoogleSheetsBatch'
    }
}
