@{
    AliasesToExport        = @('WordSection', 'WordHeader', 'WordFooter', 'WordParagraph', 'WordText', 'WordBold', 'WordItalic', 'WordList', 'WordListItem', 'WordTable', 'WordTableCondition', 'WordImage', 'WordPageNumber', 'ExcelSheet', 'ExcelCell', 'ExcelTable')
    Author                 = 'Przemyslaw Klys'
    CmdletsToExport        = @('Add-OfficeWordFooter', 'Add-OfficeWordHeader', 'Add-OfficeWordImage', 'Add-OfficeWordList', 'Add-OfficeWordListItem', 'Add-OfficeWordPageNumber', 'Add-OfficeWordParagraph', 'Add-OfficeWordSection', 'Add-OfficeWordTable', 'Add-OfficeWordTableCondition', 'Add-OfficeWordText', 'Close-OfficeWord', 'Get-OfficeWord', 'New-OfficeWord', 'Add-OfficePowerPointSlide', 'Add-OfficePowerPointTextBox', 'Get-OfficePowerPoint', 'Get-OfficePowerPointSlide', 'Merge-OfficePowerPoint', 'New-OfficePowerPoint', 'Remove-OfficePowerPointSlide', 'Save-OfficePowerPoint', 'Set-OfficePowerPointSlideTitle', 'Add-OfficeExcelSheet', 'Add-OfficeExcelTable', 'Close-OfficeExcel', 'Get-OfficeExcel', 'New-OfficeExcel', 'Set-OfficeExcelCell')
    CompanyName            = 'Evotec'
    CompatiblePSEditions   = @('Desktop', 'Core')
    Copyright              = '(c) 2011 - 2025 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description            = 'Experimental PowerShell Module to create and edit Microsoft Word, Microsoft Excel, and Microsoft PowerPoint documents without having Microsoft Office installed.'
    DotNetFrameworkVersion = '4.7.2'
    FunctionsToExport      = @()
    GUID                   = 'd75a279d-30c2-4c2d-ae0d-12f1f3bf4d39'
    ModuleVersion          = '0.3.0'
    PowerShellVersion      = '5.1'
    PrivateData            = @{
        PSData = @{
            IconUri    = 'https://evotec.xyz/wp-content/uploads/2018/10/PSWriteWord.png'
            LicenseUri = 'https://github.com/EvotecIT/PSWriteOffice/blob/master/License'
            ProjectUri = 'https://github.com/EvotecIT/PSWriteOffice'
            Tags       = @('word', 'docx', 'write', 'PSWord', 'office', 'windows', 'doc', 'pswriteword', 'linux', 'macos')
        }
    }
    RootModule             = 'PSWriteOffice.psm1'
}
