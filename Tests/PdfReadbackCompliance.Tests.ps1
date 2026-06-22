BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'PDF readback and compliance cmdlets' {
    It 'converts PDF logical readback to Markdown' {
        $pdfPath = Join-Path $TestDrive 'markdown-source.pdf'
        $markdownPath = Join-Path $TestDrive 'markdown-output.md'

        New-OfficePdf -Path $pdfPath {
            PdfHeading 'Markdown Source'
            PdfParagraph 'Logical readback should preserve this sentence.'
        } | Out-Null

        ConvertTo-OfficePdfMarkdown -Path $pdfPath -OutputPath $markdownPath | Should -BeOfType System.IO.FileInfo

        $markdown = Get-Content -Raw -Path $markdownPath
        $markdown | Should -Match 'Markdown Source'
        $markdown | Should -Match 'Logical readback'
    }

    It 'extracts image resources from generated PDFs' {
        $pdfPath = Join-Path $TestDrive 'images.pdf'
        $imagePath = Join-Path $PSScriptRoot 'Assets\CellImage.png'
        $outputDirectory = Join-Path $TestDrive 'images'

        New-OfficePdf -Path $pdfPath {
            PdfHeading 'Image Report'
            PdfImage -Path $imagePath -Width 64 -Height 64 -AlternativeText 'Sample image'
        } | Out-Null

        $images = @(Get-OfficePdfImage -Path $pdfPath)
        $images.Count | Should -BeGreaterThan 0
        $images[0].Width | Should -BeGreaterThan 0

        $files = @(Get-OfficePdfImage -Path $pdfPath -OutputDirectory $outputDirectory -BaseName 'asset')
        $files.Count | Should -BeGreaterThan 0
        Test-Path $files[0].FullName | Should -BeTrue
    }

    It 'embeds and extracts PDF attachments' {
        $pdfPath = Join-Path $TestDrive 'attachments.pdf'
        $attachmentPath = Join-Path $TestDrive 'payload.txt'
        $outputDirectory = Join-Path $TestDrive 'attachments'
        Set-Content -Path $attachmentPath -Value 'Attachment payload' -NoNewline

        New-OfficePdf -Path $pdfPath {
            PdfHeading 'Attachment Report'
            PdfAttachment -Path $attachmentPath -Name 'payload.txt' -MimeType 'text/plain' -Relationship Data -Description 'Test payload'
        } | Out-Null

        $attachments = @(Get-OfficePdfAttachment -Path $pdfPath)
        $attachments.Count | Should -Be 1
        $attachments[0].FileName | Should -Be 'payload.txt'
        $attachments[0].Relationship | Should -Be 'Data'

        $files = @(Get-OfficePdfAttachment -Path $pdfPath -OutputDirectory $outputDirectory)
        $files.Count | Should -Be 1
        Get-Content -Raw -Path $files[0].FullName | Should -Be 'Attachment payload'
    }

    It 'reports generated PDF compliance readiness' {
        $document = New-OfficePdf {
            PdfMetadata -Title 'Compliance Draft'
            PdfCompliance -Profile PdfUa1 -Groundwork
            PdfHeading 'Compliance Draft'
            PdfParagraph 'This document can be assessed before saving.'
        }

        $report = $document | Get-OfficePdfCompliance -Profile PdfUa1
        $report.Profile | Should -Be 'PdfUa1'
        $report.DisplayName | Should -BeLike '*PDF/UA*'
        $report.Requirements.Count | Should -BeGreaterThan 0
    }

    It 'reports saved PDF compliance readback evidence by path' {
        $pdfPath = Join-Path $TestDrive 'pdfua-readback.pdf'
        $options = [OfficeIMO.Pdf.PdfOptions]::new()
        $options.FileVersion = [OfficeIMO.Pdf.PdfFileVersion]::Pdf17
        $options.IncludeStandardFontToUnicodeMaps = $true

        $document = [OfficeIMO.Pdf.PdfDocument]::Create($options)
        $document.ConfigurePdfUaGroundwork('en-US') | Out-Null
        $document.Meta('Readback PDF/UA', 'PSWriteOffice', $null, $null) | Out-Null
        $document.H1('Readback PDF/UA') | Out-Null
        $document.Paragraph({ param($p) $p.Text('Saved PDF compliance readback evidence') }) | Out-Null
        $document.Save($pdfPath)

        $report = Get-OfficePdfCompliance -Path $pdfPath -Profile PdfUa1

        $report.Profile | Should -Be 'PdfUa1'
        $report.FindRequirement('readback-pdfua-identification').Status | Should -Be 'Satisfied'
        $report.FindRequirement('readback-document-title').Status | Should -Be 'Satisfied'
        $report.FindRequirement('readback-marked-catalog').Status | Should -Be 'Satisfied'
        $report.FindRequirement('pdfua-validation').Status | Should -Be 'Unsupported'
    }

    It 'configures e-invoice associated file groundwork from XML' {
        $xmlPath = Join-Path $TestDrive 'invoice.xml'
        $pdfPath = Join-Path $TestDrive 'einvoice.pdf'
        @'
<?xml version="1.0" encoding="UTF-8"?>
<rsm:CrossIndustryInvoice xmlns:rsm="urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100" xmlns:ram="urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100">
  <rsm:ExchangedDocumentContext>
    <ram:GuidelineSpecifiedDocumentContextParameter>
      <ram:ID>urn:factur-x.eu:1p0:en16931</ram:ID>
    </ram:GuidelineSpecifiedDocumentContextParameter>
  </rsm:ExchangedDocumentContext>
</rsm:CrossIndustryInvoice>
'@ | Set-Content -Path $xmlPath -Encoding UTF8

        New-OfficePdf -Path $pdfPath {
            PdfElectronicInvoice -Path $xmlPath -Profile FacturX -ConformanceLevel 'EN 16931'
            PdfMetadata -Title 'E-invoice'
            PdfHeading 'E-invoice'
            PdfParagraph 'The CII XML payload is embedded as an associated file.'
        } | Out-Null

        $attachment = Get-OfficePdfAttachment -Path $pdfPath
        $report = Get-OfficePdfCompliance -Path $pdfPath -Profile FacturX

        $attachment.FileName | Should -Be 'factur-x.xml'
        $attachment.MimeType | Should -Be 'application/xml'
        $attachment.Relationship | Should -Be 'Data'
        $report.FindRequirement('readback-einvoice-xmp').Status | Should -Be 'Satisfied'
        $report.FindRequirement('readback-associated-invoice-file').Status | Should -Be 'Satisfied'
    }
}
