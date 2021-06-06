@{
    AliasesToExport      = @()
    Author               = 'Przemyslaw Klys'
    CmdletsToExport      = @()
    CompanyName          = 'Evotec'
    CompatiblePSEditions = @('Desktop', 'Core')
    Copyright            = '(c) 2011 - 2021 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description          = 'Experimental PowerShell Module to create and edit Microsoft Word, Microsoft Excel, and Microsoft PowerPoint documents without having Microsoft Office installed.'
    FunctionsToExport    = @('Close-OfficeWord', 'ConvertFrom-HTMLtoWord', 'Export-OfficeExcel', 'Get-OfficeExcel', 'Get-OfficeExcelValue', 'Get-OfficeExcelWorkSheet', 'Get-OfficeWord', 'New-OfficeExcel', 'New-OfficeExcelTable', 'New-OfficeExcelTableOptions', 'New-OfficeExcelValue', 'New-OfficeExcelWorkSheet', 'New-OfficePowerPoint', 'New-OfficeWord', 'New-OfficeWordTable', 'New-OfficeWordTableBorder', 'New-OfficeWordTableLayout', 'New-OfficeWordTableLook', 'New-OfficeWordTableStyle', 'New-OfficeWordText', 'Save-OfficeExcel', 'Save-OfficePowerPoint', 'Save-OfficeWord', 'Set-OfficeExcelCellStyle', 'Set-OfficeExcelWorkSheetStyle')
    GUID                 = 'd75a279d-30c2-4c2d-ae0d-12f1f3bf4d39'
    ModuleVersion        = '0.0.1'
    PowerShellVersion    = '5.1'
    PrivateData          = @{
        PSData = @{
            Tags       = @('word', 'docx', 'write', 'PSWord', 'office', 'windows', 'doc', 'pswriteword', 'linux', 'macos')
            LicenseUri = 'https://github.com/EvotecIT/PSWriteOffice/blob/master/License'
            ProjectUri = 'https://github.com/EvotecIT/PSWriteOffice'
            IconUri    = 'https://evotec.xyz/wp-content/uploads/2018/10/PSWriteWord.png'
        }
    }
    RequiredModules      = @(@{
            ModuleVersion = '0.0.205'
            ModuleName    = 'PSSharedGoods'
            Guid          = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
        })
    RootModule           = 'PSWriteOffice.psm1'
}