@{
    AliasesToExport        = @()
    Author                 = 'Przemyslaw Klys'
    CmdletsToExport        = @('Close-OfficeWord', 'ConvertFrom-HTMLtoWord', 'Get-OfficeWord', 'New-OfficeWord', 'New-OfficeWordList', 'New-OfficeWordListItem', 'New-OfficeWordTable', 'New-OfficeWordText', 'Remove-OfficeWordFooter', 'Remove-OfficeWordHeader', 'Save-OfficeWord', 'Add-OfficePowerPointSlide', 'Add-OfficePowerPointTextBox', 'Get-OfficePowerPoint', 'Get-OfficePowerPointSlide', 'Merge-OfficePowerPoint', 'New-OfficePowerPoint', 'Remove-OfficePowerPointSlide', 'Save-OfficePowerPoint', 'Set-OfficePowerPointSlideTitle', 'Close-OfficeExcel', 'Export-OfficeExcel', 'Get-OfficeExcel', 'Get-OfficeExcelValue', 'Get-OfficeExcelWorkSheet', 'Get-OfficeExcelWorkSheetData', 'Import-OfficeExcel', 'New-OfficeExcel', 'New-OfficeExcelTable', 'New-OfficeExcelValue', 'New-OfficeExcelWorkSheet', 'Save-OfficeExcel', 'Set-OfficeExcelCellStyle', 'Set-OfficeExcelWorkSheetStyle')
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