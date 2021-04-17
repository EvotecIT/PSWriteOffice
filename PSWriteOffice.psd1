@{
    AliasesToExport      = @()
    Author               = 'Przemyslaw Klys'
    CmdletsToExport      = @()
    CompanyName          = 'Evotec'
    CompatiblePSEditions = @('Desktop', 'Core')
    Copyright            = '(c) 2011 - 2021 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description          = 'Experimental PowerShell Module to create and edit Microsoft Word, Microsoft Excel, and Microsoft PowerPoint documents without having Microsoft Office installed.'
    FunctionsToExport    = @('Add-OfficeExcelWorkSheet', 'Get-OfficeExcel', 'Get-OfficeExcelWorkSheet', 'New-OfficeExcel', 'New-OfficeExcelTable', 'New-OfficePowerPoint', 'New-OfficeWord', 'New-OfficeWordText', 'Save-OfficeExcel', 'Save-OfficePowerPoint', 'Save-OfficeWord')
    GUID                 = 'd75a279d-30c2-4c2d-ae0d-12f1f3bf4d39'
    ModuleVersion        = '0.0.1'
    PowerShellVersion    = '5.1'
    PrivateData          = @{
        PSData = @{
            Tags       = @('word', 'docx', 'write', 'PSWord', 'office', 'windows', 'doc')
            ProjectUri = 'https://github.com/EvotecIT/DocumentoZaurr'
        }
    }
    RequiredModules      = @(@{
            ModuleName    = 'PSSharedGoods'
            ModuleVersion = '0.0.199'
            Guid          = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
        })
    RootModule           = 'PSWriteOffice.psm1'
}