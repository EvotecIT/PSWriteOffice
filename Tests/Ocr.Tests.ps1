BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    function New-TestTesseract {
        param([Parameter(Mandatory)][string] $Directory)

        $onWindows = $PSVersionTable.PSEdition -eq 'Desktop' -or $IsWindows
        if ($onWindows) {
            $path = Join-Path $Directory 'fake-tesseract.cmd'
            @'
@echo off
if "%~1"=="--version" (
  echo tesseract 5.5.0
  exit /b 0
)
if "%~1"=="--list-langs" (
  echo List of available languages in fake runtime:
  echo eng
  echo pol
  exit /b 0
)
set "output=%~2.tsv"
>"%output%" echo level	page_num	block_num	par_num	line_num	word_num	left	top	width	height	conf	text
>>"%output%" echo 5	1	1	1	1	1	10	12	120	24	96.0	OfficeIMO
>>"%output%" echo 5	1	1	1	1	2	140	12	72	24	94.0	OCR
exit /b 0
'@ | Set-Content -LiteralPath $path -Encoding Ascii
            return $path
        }

        $path = Join-Path $Directory 'fake-tesseract'
        @'
#!/usr/bin/env sh
if [ "$1" = "--version" ]; then
  echo "tesseract 5.5.0"
  exit 0
fi
if [ "$1" = "--list-langs" ]; then
  echo "List of available languages in fake runtime:"
  echo "eng"
  echo "pol"
  exit 0
fi
printf 'level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext\n' > "${2}.tsv"
printf '5\t1\t1\t1\t1\t1\t10\t12\t120\t24\t96.0\tOfficeIMO\n' >> "${2}.tsv"
printf '5\t1\t1\t1\t1\t2\t140\t12\t72\t24\t94.0\tOCR\n' >> "${2}.tsv"
'@ | Set-Content -LiteralPath $path -Encoding utf8NoBOM
        & chmod '+x' $path
        return $path
    }
}

Describe 'Easy local OCR commands' {
    It 'recognizes image text with runtime and word evidence' {
        $runtime = New-TestTesseract -Directory $TestDrive
        $result = Get-OfficeImageText `
            -Path (Join-Path $PSScriptRoot 'Assets\CellImage.png') `
            -TesseractPath $runtime `
            -Language eng+pol `
            -NoLanguageDownload `
            -PassThru

        $result.Text | Should -Be 'OfficeIMO OCR'
        $result.Provider | Should -Be 'tesseract-cli'
        $result.Model | Should -Be 'tessdata:eng+pol'
        @($result.Spans).Count | Should -BeGreaterThan 0
    }

    It 'writes a searchable PDF whose OCR text is readable' {
        $runtime = New-TestTesseract -Directory $TestDrive
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
        $runtime = New-TestTesseract -Directory $TestDrive
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
