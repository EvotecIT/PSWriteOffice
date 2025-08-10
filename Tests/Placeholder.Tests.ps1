Describe 'Placeholder classes' {
    It 'GetOfficeWordCommand exists' {
        [PSWriteOffice.Cmdlets.Word.GetOfficeWordCommand] | Should -Not -BeNullOrEmpty
    }
    It 'GetOfficeExcelCommand exists' {
        [PSWriteOffice.Cmdlets.Excel.GetOfficeExcelCommand] | Should -Not -BeNullOrEmpty
    }
    It 'GetOfficePowerPointCommand exists' {
        [PSWriteOffice.Cmdlets.PowerPoint.GetOfficePowerPointCommand] | Should -Not -BeNullOrEmpty
    }
    It 'WordDocumentService exists' {
        [PSWriteOffice.Services.Word.WordDocumentService] | Should -Not -BeNullOrEmpty
    }
    It 'ExcelDocumentService exists' {
        [PSWriteOffice.Services.Excel.ExcelDocumentService] | Should -Not -BeNullOrEmpty
    }
    It 'PowerPointDocumentService exists' {
        [PSWriteOffice.Services.PowerPoint.PowerPointDocumentService] | Should -Not -BeNullOrEmpty
    }
    It 'AddOfficePowerPointSlideCommand exists' {
        [PSWriteOffice.Cmdlets.PowerPoint.AddOfficePowerPointSlideCommand] | Should -Not -BeNullOrEmpty
    }
    It 'MergeOfficePowerPointCommand exists' {
        [PSWriteOffice.Cmdlets.PowerPoint.MergeOfficePowerPointCommand] | Should -Not -BeNullOrEmpty
    }
}
