BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Easy local OCR commands' {
    It 'exposes discoverable typed language choices' {
        $parameter = (Get-Command Get-OfficeImageText).Parameters['Language']
        $parameter.ParameterType.FullName | Should -Be 'OfficeIMO.Ocr.Tesseract.TesseractOcrLanguage[]'
        $languageType = $parameter.ParameterType.GetElementType()
        $languageType.GetEnumNames() | Should -Contain 'English'
        $languageType.GetEnumNames() | Should -Contain 'Polish'
        (Get-Command New-OfficeDocumentReader).Parameters['OcrLanguage'].ParameterType.FullName |
            Should -Be 'OfficeIMO.Ocr.Tesseract.TesseractOcrLanguage[]'
    }

    It 'recognizes image text with runtime and word evidence' {
        $runtime = New-TestTesseractExecutable -Directory $TestDrive
        $result = Get-OfficeImageText `
            -Path (Join-Path $PSScriptRoot 'Assets\CellImage.png') `
            -TesseractPath $runtime `
            -Language English, Polish `
            -NoLanguageDownload `
            -PassThru

        $result.Text | Should -Be 'OfficeIMO OCR'
        $result.Provider | Should -Be 'tesseract-cli'
        $result.Model | Should -Be 'tessdata:eng+pol'
        @($result.Spans).Count | Should -BeGreaterThan 0
    }

    It 'rejects unsupported image formats before discovering an OCR runtime' {
        $unsupportedPath = Join-Path $TestDrive 'not-an-image.txt'
        [IO.File]::WriteAllText($unsupportedPath, 'This is not an OCR image.')

        {
            Get-OfficeImageText `
                -Path $unsupportedPath `
                -TesseractPath 'definitely-missing-tesseract-for-pswriteoffice' `
                -NoLanguageDownload `
                -ErrorAction Stop
        } | Should -Throw '*supports PNG, JPEG, TIFF, BMP, GIF, WebP, and JPEG 2000*'
    }

    It 'preserves session options and lets friendly language parameters override them' {
        $runtime = New-TestTesseractExecutable -Directory $TestDrive
        $optionsType = (Get-Command Get-OfficeImageText).Parameters['Options'].ParameterType
        $options = [Activator]::CreateInstance($optionsType)
        $options.Engine.Language = 'pol'

        $legacyResult = Get-OfficeImageText `
            -Path (Join-Path $PSScriptRoot 'Assets\CellImage.png') `
            -Options $options `
            -TesseractPath $runtime `
            -NoLanguageDownload `
            -PassThru
        $legacyResult.Model | Should -Be 'tessdata:pol'

        $friendlyResult = Get-OfficeImageText `
            -Path (Join-Path $PSScriptRoot 'Assets\CellImage.png') `
            -Options $options `
            -TesseractPath $runtime `
            -Language English, Polish `
            -NoLanguageDownload `
            -PassThru
        $friendlyResult.Model | Should -Be 'tessdata:eng+pol'

        $rawResult = Get-OfficeImageText `
            -Path (Join-Path $PSScriptRoot 'Assets\CellImage.png') `
            -Options $options `
            -TesseractPath $runtime `
            -TesseractLanguageExpression 'eng+pol' `
            -NoLanguageDownload `
            -PassThru
        $rawResult.Model | Should -Be 'tessdata:eng+pol'

        {
            Get-OfficeImageText `
                -Path (Join-Path $PSScriptRoot 'Assets\CellImage.png') `
                -TesseractPath $runtime `
                -Language English `
                -TesseractLanguageExpression 'eng+pol' `
                -NoLanguageDownload
        } | Should -Throw
    }

    It 'writes a searchable PDF whose OCR text is readable' {
        $runtime = New-TestTesseractExecutable -Directory $TestDrive
        $inputPath = Join-Path $TestDrive 'scan.pdf'
        $outputPath = Join-Path $TestDrive 'scan-searchable.pdf'
        $imagePath = Join-Path $PSScriptRoot 'Assets\CellImage.png'

        New-OfficePdf -Path $inputPath {
            PdfImage -Path $imagePath -Width 180 -Height 120
        }

        $result = ConvertTo-OfficePdfSearchable `
            -Path $inputPath `
            -OutputPath $outputPath `
            -TesseractPath $runtime `
            -NoLanguageDownload `
            -PassThru

        $result.WasModified | Should -BeTrue
        $result.AddedWordCount | Should -Be 2
        Test-Path -LiteralPath $outputPath | Should -BeTrue
        Get-OfficePdfText -Path $outputPath | Should -Match 'OfficeIMO OCR'
    }

    It 'preserves an existing searchable PDF destination unless Force is supplied' {
        $runtime = New-TestTesseractExecutable -Directory $TestDrive
        $inputPath = Join-Path $TestDrive 'force-input.pdf'
        $outputPath = Join-Path $TestDrive 'force-output.pdf'
        $imagePath = Join-Path $PSScriptRoot 'Assets\CellImage.png'

        New-OfficePdf -Path $inputPath {
            PdfImage -Path $imagePath -Width 180 -Height 120
        }
        [IO.File]::WriteAllText($outputPath, 'caller-owned')

        {
            ConvertTo-OfficePdfSearchable `
                -Path $inputPath `
                -OutputPath $outputPath `
                -TesseractPath $runtime `
                -NoLanguageDownload
        } | Should -Throw
        [IO.File]::ReadAllText($outputPath) | Should -Be 'caller-owned'

        ConvertTo-OfficePdfSearchable `
            -Path $inputPath `
            -OutputPath $outputPath `
            -TesseractPath $runtime `
            -NoLanguageDownload `
            -Force
        Get-OfficePdfText -Path $outputPath | Should -Match 'OfficeIMO OCR'
    }
}
