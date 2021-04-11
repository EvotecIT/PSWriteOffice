@{
    AliasesToExport      = @()
    Author               = 'Przemyslaw Klys'
    CmdletsToExport      = @()
    CompanyName          = 'Evotec'
    CompatiblePSEditions = @('Desktop', 'Core')
    Copyright            = '(c) 2011 - 2021 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description          = 'Simple project to create Microsoft Word in PowerShell without having Office installed.'
    FunctionsToExport    = @('Add-OfficeExcelWorkSheet', 'New-OfficeExcel', 'New-OfficePowerPoint', 'New-OfficeWord', 'New-OfficeWordText', 'Save-OfficeExcel', 'Save-OfficePowerPoint', 'Save-OfficeWord')
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
            ModuleVersion = '0.0.198'
            Guid          = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
            ModuleName    = 'PSSharedGoods'
        })
    RootModule           = 'DocumentoZaurr.psm1'
}