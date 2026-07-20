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
            'Get-OfficeDocumentBatch', 'New-OfficeDocumentReader', 'Search-OfficeDocument',
            'Get-OfficeDocumentPageMarkdown',
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
        (Get-Command Import-OfficePdfXfdf).Parameters.Keys | Should -Contain 'MaxXfdfBytes'
        (Get-Command Import-OfficePdfXfdf).Parameters.Keys | Should -Contain 'ReadOptions'
        (Get-Command Export-OfficePdfXfdf).Parameters.Keys | Should -Contain 'ReadOptions'
        (Get-Command ConvertTo-OfficePdfSanitized).Parameters.Keys | Should -Contain 'ReadOptions'
        (Get-Command Export-OfficePdfImage).Parameters.Options.ParameterType.FullName |
            Should -Be 'OfficeIMO.Pdf.PdfImageExportOptions'
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

        $brokenAsciiDocPath = Join-Path $TestDrive 'broken.adoc'
        $brokenAsciiDocOutput = Join-Path $TestDrive 'broken-adoc.md'
        Set-Content -LiteralPath $brokenAsciiDocPath -Encoding UTF8 -Value "----`nunterminated block"
        { ConvertTo-OfficeAsciiDocMarkdown -Path $brokenAsciiDocPath -OutputPath $brokenAsciiDocOutput -FailOnLoss -ErrorAction Stop } |
            Should -Throw '*parsing reported errors*'
        Test-Path -LiteralPath $brokenAsciiDocOutput | Should -BeFalse

        $brokenLatexPath = Join-Path $TestDrive 'broken.tex'
        $brokenLatexOutput = Join-Path $TestDrive 'broken-latex.md'
        Set-Content -LiteralPath $brokenLatexPath -Encoding UTF8 -Value '\textbf{unterminated'
        { ConvertTo-OfficeLatexMarkdown -Path $brokenLatexPath -OutputPath $brokenLatexOutput -FailOnLoss -ErrorAction Stop } |
            Should -Throw '*parsing reported errors*'
        Test-Path -LiteralPath $brokenLatexOutput | Should -BeFalse
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

    It 'protects OpenDocument destinations and validates conversion extensions' {
        $whatIfPath = Join-Path $TestDrive 'what-if.odt'
        New-OfficeOpenDocument -Kind Text -Path $whatIfPath -WhatIf | Out-Null
        Test-Path -LiteralPath $whatIfPath | Should -BeFalse

        $signedPath = Join-Path $TestDrive 'signed.odt'
        New-OfficeOpenDocument -Kind Text -Path $signedPath | Out-Null
        $stream = [System.IO.File]::Open($signedPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
        try {
            $archive = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Update, $false)
            try {
                $entry = $archive.CreateEntry('META-INF/documentsignatures.xml')
                $writer = [System.IO.StreamWriter]::new($entry.Open())
                try { $writer.Write('<?xml version="1.0"?><signatures/>') } finally { $writer.Dispose() }
            } finally {
                $archive.Dispose()
            }
        } finally {
            $stream.Dispose()
        }

        $signed = Get-OfficeOpenDocument -Path $signedPath
        $signed.Metadata.Title = 'Invalidates the preserved signature'
        $odfSaveOptionsType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.OpenDocument' -TypeName 'OfficeIMO.OpenDocument.OdfSaveOptions' -CommandName 'Save-OfficeOpenDocument'
        $saveOptions = [Activator]::CreateInstance($odfSaveOptionsType)
        $saveOptions.SignatureHandling = 'RemoveInvalidated'
        $rejectedPath = Join-Path $TestDrive 'must-not-exist.odt'
        { $signed | Save-OfficeOpenDocument -Path $rejectedPath -Options $saveOptions -FailOnLoss -ErrorAction Stop } |
            Should -Throw
        Test-Path -LiteralPath $rejectedPath | Should -BeFalse

        $excelPath = Join-Path $TestDrive 'extension-source.xlsx'
        $wrongOutput = Join-Path $TestDrive 'spreadsheet.odt'
        New-OfficeExcel -Path $excelPath { ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Extension contract' } } | Out-Null
        { ConvertTo-OfficeOpenDocument -Path $excelPath -OutputPath $wrongOutput -ErrorAction Stop } |
            Should -Throw '*must use the .ods extension*'
        Test-Path -LiteralPath $wrongOutput | Should -BeFalse

        $textPath = Join-Path $TestDrive 'reverse-source.odt'
        $wrongOfficeOutput = Join-Path $TestDrive 'text.xlsx'
        New-OfficeOpenDocument -Kind Text -Path $textPath | Out-Null
        { ConvertFrom-OfficeOpenDocument -Path $textPath -OutputPath $wrongOfficeOutput -ErrorAction Stop } |
            Should -Throw '*must use the .docx extension*'
        Test-Path -LiteralPath $wrongOfficeOutput | Should -BeFalse

        $wrongNewPath = Join-Path $TestDrive 'new-spreadsheet.odt'
        { New-OfficeOpenDocument -Kind Spreadsheet -Path $wrongNewPath -ErrorAction Stop } |
            Should -Throw '*must use the .ods extension*'
        Test-Path -LiteralPath $wrongNewPath | Should -BeFalse

        $spreadsheet = New-OfficeOpenDocument -Kind Spreadsheet
        $wrongSavePath = Join-Path $TestDrive 'saved-spreadsheet.odt'
        { $spreadsheet | Save-OfficeOpenDocument -Path $wrongSavePath -ErrorAction Stop } |
            Should -Throw '*must use the .ods extension*'
        Test-Path -LiteralPath $wrongSavePath | Should -BeFalse
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
            Export-OfficePdfImage -Path $pdfPath -OutputPath (Join-Path $TestDrive 'pdf-webp-images') -Format Webp
        )

        $results.Count | Should -BeGreaterOrEqual 5
        foreach ($result in $results) {
            $result.GetType().FullName | Should -Be 'OfficeIMO.Drawing.OfficeImageExportResult'
            $result.Bytes.Length | Should -BeGreaterThan 0
        }
        $webp = $results | Where-Object Format -EQ Webp | Select-Object -First 1
        $webp | Should -Not -BeNullOrEmpty
        Test-Path -LiteralPath $webp.SavedPath | Should -BeTrue

        $htmlLines = @('<html><body>', '<h1>Pipeline heading</h1><p>Pipeline body</p>', '</body></html>')
        $pipelineHtmlPath = Join-Path $TestDrive 'pipeline-html.svg'
        $pipelineHtml = @($htmlLines | Export-OfficeHtmlImage -OutputPath $pipelineHtmlPath -Format Svg)
        $pipelineHtml | Should -HaveCount 1
        $pipelineHtml[0].GetType().FullName | Should -Be 'OfficeIMO.Drawing.OfficeImageExportResult'
        Test-Path -LiteralPath $pipelineHtmlPath | Should -BeTrue
    }

    It 'releases path-loaded PowerPoint presentations after image export' {
        $powerPointPath = Join-Path $TestDrive 'transient.pptx'
        New-OfficePowerPoint -Path $powerPointPath { PptSlide { PptTitle -Title 'Transient' } } | Out-Null
        $serviceType = Get-TestPSWriteOfficeType -AssemblyName 'PSWriteOffice' -TypeName 'PSWriteOffice.Services.PowerPoint.PowerPointDocumentService' -CommandName 'Export-OfficePowerPointImage'
        $presentationsField = $serviceType.GetField('Presentations', [System.Reflection.BindingFlags]'NonPublic, Static')
        $presentations = $presentationsField.GetValue($null)
        $before = $presentations.Count

        Export-OfficePowerPointImage -Path $powerPointPath -OutputPath (Join-Path $TestDrive 'transient-images') -Format Svg | Out-Null
        $presentations.Count | Should -Be $before

        Get-OfficePowerPointInspection -Path $powerPointPath | Out-Null

        $presentations.Count | Should -Be $before
    }

    It 'releases path-loaded Office documents after OpenDocument conversion' {
        $flags = [System.Reflection.BindingFlags]'NonPublic, Static'
        $wordServiceType = Get-TestPSWriteOfficeType -AssemblyName 'PSWriteOffice' -TypeName 'PSWriteOffice.Services.Word.WordDocumentService' -CommandName 'ConvertTo-OfficeOpenDocument'
        $excelServiceType = Get-TestPSWriteOfficeType -AssemblyName 'PSWriteOffice' -TypeName 'PSWriteOffice.Services.Excel.ExcelDocumentService' -CommandName 'ConvertTo-OfficeOpenDocument'
        $powerPointServiceType = Get-TestPSWriteOfficeType -AssemblyName 'PSWriteOffice' -TypeName 'PSWriteOffice.Services.PowerPoint.PowerPointDocumentService' -CommandName 'ConvertTo-OfficeOpenDocument'
        $wordDocuments = $wordServiceType.GetField('AssociatedPaths', $flags).GetValue($null)
        $excelDocuments = $excelServiceType.GetField('AssociatedPaths', $flags).GetValue($null)
        $presentations = $powerPointServiceType.GetField('Presentations', $flags).GetValue($null)
        $beforeWord = $wordDocuments.Count
        $beforeExcel = $excelDocuments.Count
        $beforePowerPoint = $presentations.Count

        $wordPath = Join-Path $TestDrive 'convert.docx'
        $excelPath = Join-Path $TestDrive 'convert.xlsx'
        $powerPointPath = Join-Path $TestDrive 'convert.pptx'
        New-OfficeWord -Path $wordPath { WordSection { WordParagraph -Text 'Convert' } } | Out-Null
        New-OfficeExcel -Path $excelPath { ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Convert' } } | Out-Null
        New-OfficePowerPoint -Path $powerPointPath { PptSlide { PptTitle -Title 'Convert' } } | Out-Null

        ConvertTo-OfficeOpenDocument -Path $wordPath -OutputPath (Join-Path $TestDrive 'convert.odt') | Out-Null
        ConvertTo-OfficeOpenDocument -Path $excelPath -OutputPath (Join-Path $TestDrive 'convert.ods') | Out-Null
        ConvertTo-OfficeOpenDocument -Path $powerPointPath -OutputPath (Join-Path $TestDrive 'convert.odp') | Out-Null

        $wordDocuments.Count | Should -Be $beforeWord
        $excelDocuments.Count | Should -Be $beforeExcel
        $presentations.Count | Should -Be $beforePowerPoint
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

        $formPdfPath = Join-Path $TestDrive 'advanced-form.pdf'
        New-OfficePdf -Path $formPdfPath {
            PdfParagraph 'Advanced PDF form proof'
            PdfFormField -Name 'Reviewer' -Type Text -Value 'Initial'
        } | Out-Null
        $xfdf = Export-OfficePdfXfdf -Path $formPdfPath
        $xfdfLines = @(($xfdf -replace '><', ">`n<") -split "`n")
        $xfdfOutput = Join-Path $TestDrive 'advanced-xfdf.pdf'
        $imported = @($xfdfLines | Import-OfficePdfXfdf -Path $formPdfPath -OutputPath $xfdfOutput -PassThru)
        $imported | Should -HaveCount 1
        $imported[0].GetType().FullName | Should -Be 'OfficeIMO.Pdf.PdfDocument'
        Test-Path -LiteralPath $xfdfOutput | Should -BeTrue

        $xfdfPath = Join-Path $TestDrive 'advanced.xfdf'
        Set-Content -LiteralPath $xfdfPath -Value $xfdf -Encoding UTF8
        $boundedOutput = Join-Path $TestDrive 'bounded-xfdf.pdf'
        { Import-OfficePdfXfdf -Path $formPdfPath -XfdfPath $xfdfPath -OutputPath $boundedOutput -MaxXfdfBytes 8 -ErrorAction Stop } |
            Should -Throw '*configured limit*'
        Test-Path -LiteralPath $boundedOutput | Should -BeFalse

        $pdfReadOptionsType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Pdf' -TypeName 'OfficeIMO.Pdf.PdfReadOptions' -CommandName 'Compare-OfficePdfVisual'
        $differenceReadOptions = [Activator]::CreateInstance($pdfReadOptionsType)
        $differenceReadOptions.Limits.MaxInputBytes = 1
        { Compare-OfficePdfVisual -ReferencePath $pdfPath -DifferencePath $pdfPath -DifferenceReadOptions $differenceReadOptions -ErrorAction Stop } |
            Should -Throw

        $boundedSanitizedOutput = Join-Path $TestDrive 'bounded-sanitized.pdf'
        { ConvertTo-OfficePdfSanitized -Path $pdfPath -OutputPath $boundedSanitizedOutput -ReadOptions $differenceReadOptions -ErrorAction Stop } |
            Should -Throw
        Test-Path -LiteralPath $boundedSanitizedOutput | Should -BeFalse

        $boundedPdfImportOutput = Join-Path $TestDrive 'bounded-source-import.pdf'
        { Import-OfficePdfXfdf -Path $formPdfPath -Xfdf $xfdf -OutputPath $boundedPdfImportOutput -ReadOptions $differenceReadOptions -ErrorAction Stop } |
            Should -Throw
        Test-Path -LiteralPath $boundedPdfImportOutput | Should -BeFalse

        { Export-OfficePdfXfdf -Path $formPdfPath -ReadOptions $differenceReadOptions -ErrorAction Stop } |
            Should -Throw

        $rewriteOptionsType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Pdf' -TypeName 'OfficeIMO.Pdf.PdfRewritePreservationOptions' -CommandName 'Test-OfficePdfRewrite'
        $rewriteOptions = [Activator]::CreateInstance($rewriteOptionsType)
        $rewriteOptions.OriginalReadOptions = $differenceReadOptions
        { Test-OfficePdfRewrite -ReferencePath $pdfPath -DifferencePath $safePath -Options $rewriteOptions -ErrorAction Stop } |
            Should -Throw
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
