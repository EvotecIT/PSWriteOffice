Describe 'Placeholder classes' {
    It 'GetOfficeWordCommand exists' {
        [Type]::GetType('PSWriteOffice.Cmdlets.Word.GetOfficeWordCommand', $false) | Should -Not -BeNullOrEmpty
    }
    It 'GetOfficeExcelCommand exists' {
        [Type]::GetType('PSWriteOffice.Cmdlets.Excel.GetOfficeExcelCommand', $false) | Should -Not -BeNullOrEmpty
    }
    It 'GetOfficePowerPointCommand exists' {
        [Type]::GetType('PSWriteOffice.Cmdlets.PowerPoint.GetOfficePowerPointCommand', $false) | Should -Not -BeNullOrEmpty
    }
    It 'WordDocumentService exists' {
        [Type]::GetType('PSWriteOffice.Services.Word.WordDocumentService', $false) | Should -Not -BeNullOrEmpty
    }
    It 'ExcelDocumentService exists' {
        [Type]::GetType('PSWriteOffice.Services.Excel.ExcelDocumentService', $false) | Should -Not -BeNullOrEmpty
    }
    It 'PowerPointDocumentService exists' {
        [Type]::GetType('PSWriteOffice.Services.PowerPoint.PowerPointDocumentService', $false) | Should -Not -BeNullOrEmpty
    }
}
