@{
    AliasesToExport        = @('WordSection', 'WordHeader', 'WordFooter', 'WordParagraph', 'WordText', 'WordBold', 'WordItalic', 'WordList', 'WordListItem', 'WordTable', 'WordTableCondition', 'WordImage', 'WordField', 'WordPageNumber', 'WordWatermark', 'Convert-WordToHtml', 'Convert-HtmlToWord', 'Convert-MarkdownToHtml', 'ExcelSheet', 'ExcelCell', 'ExcelRow', 'ExcelColumn', 'ExcelTable', 'ExcelNamedRange', 'ExcelFormula', 'ExcelHeaderFooter', 'ExcelAutoFit', 'ExcelValidationList')
    Author                 = 'Przemyslaw Klys'
    CmdletsToExport        = @('Add-OfficeWordFooter', 'Add-OfficeWordHeader', 'Add-OfficeWordImage', 'Add-OfficeWordList', 'Add-OfficeWordListItem', 'Add-OfficeWordPageNumber', 'Add-OfficeWordParagraph', 'Add-OfficeWordSection', 'Add-OfficeWordTable', 'Add-OfficeWordTableCondition', 'Add-OfficeWordText', 'Add-OfficeWordField', 'Add-OfficeWordWatermark', 'Protect-OfficeWordDocument', 'Close-OfficeWord', 'Save-OfficeWord', 'Get-OfficeWord', 'Get-OfficeWordSection', 'Get-OfficeWordParagraph', 'Get-OfficeWordTable', 'Get-OfficeWordRun', 'New-OfficeWord', 'Find-OfficeWord', 'Get-OfficeWordBookmark', 'Get-OfficeWordField', 'ConvertTo-OfficeWordHtml', 'ConvertFrom-OfficeWordHtml', 'Add-OfficePowerPointSlide', 'Add-OfficePowerPointTextBox', 'Add-OfficePowerPointShape', 'Add-OfficePowerPointTable', 'Get-OfficePowerPoint', 'Get-OfficePowerPointSlide', 'New-OfficePowerPoint', 'Remove-OfficePowerPointSlide', 'Save-OfficePowerPoint', 'Set-OfficePowerPointSlideTitle', 'Add-OfficeExcelSheet', 'Add-OfficeExcelTable', 'Add-OfficeExcelValidationList', 'Close-OfficeExcel', 'Save-OfficeExcel', 'Get-OfficeExcel', 'Get-OfficeExcelData', 'Get-OfficeExcelNamedRange', 'Get-OfficeExcelTable', 'Invoke-OfficeExcelAutoFit', 'New-OfficeExcel', 'Set-OfficeExcelCell', 'Set-OfficeExcelRow', 'Set-OfficeExcelColumn', 'Set-OfficeExcelNamedRange', 'Set-OfficeExcelFormula', 'Set-OfficeExcelHeaderFooter', 'Get-OfficeMarkdown', 'ConvertTo-OfficeMarkdown', 'ConvertTo-OfficeMarkdownHtml', 'Get-OfficeCsv', 'Get-OfficeCsvData', 'ConvertTo-OfficeCsv')
    CompanyName            = 'Evotec'
    CompatiblePSEditions   = @('Desktop', 'Core')
    Copyright              = '(c) 2011 - 2025 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description            = 'PowerShell module to create and read Microsoft Word, Excel, PowerPoint (experimental), Markdown, and CSV documents without Microsoft Office installed. Powered by OfficeIMO.*.'
    DotNetFrameworkVersion = '4.7.2'
    FunctionsToExport      = @()
    GUID                   = 'd75a279d-30c2-4c2d-ae0d-12f1f3bf4d39'
    ModuleVersion          = '0.3.0'
    PowerShellVersion      = '5.1'
    PrivateData            = @{
        PSData = @{
            LicenseUri = 'https://github.com/EvotecIT/PSWriteOffice/blob/master/License'
            ProjectUri = 'https://github.com/EvotecIT/PSWriteOffice'
            Tags       = @('officeimo', 'word', 'excel', 'powerpoint', 'markdown', 'csv', 'docx', 'xlsx', 'pptx', 'openxml', 'windows', 'linux', 'macos')
        }
    }
    RootModule             = 'PSWriteOffice.psm1'
}
