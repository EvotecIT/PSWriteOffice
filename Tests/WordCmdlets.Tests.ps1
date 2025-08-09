Describe 'Word cmdlets' {
    It 'creates new word document' {
        $path = Join-Path $TestDrive 'test.docx'
        $doc = New-OfficeWord -FilePath $path
        $doc | Should -Not -BeNullOrEmpty
        Close-OfficeWord -Document $doc
    }

    It 'saves document to path' {
        $path = Join-Path $TestDrive 'save.docx'
        $doc = New-OfficeWord -FilePath $path
        Save-OfficeWord -Document $doc
        Test-Path $path | Should -BeTrue
    }

    It 'adds text and returns paragraph' {
        $path = Join-Path $TestDrive 'text.docx'
        $doc = New-OfficeWord -FilePath $path
        $para = New-OfficeWordText -Document $doc -Text 'hello' -ReturnObject
        $para.Text | Should -Be 'hello'
        Close-OfficeWord -Document $doc
    }

    It 'throws when optional array length mismatches Text' {
        $path = Join-Path $TestDrive 'mismatch.docx'
        $doc = New-OfficeWord -FilePath $path
        { New-OfficeWordText -Document $doc -Text @('one','two') -Bold @($true) } | Should -Throw
        { New-OfficeWordText -Document $doc -Text @('one','two') -Color @('FF0000') } | Should -Throw
        Close-OfficeWord -Document $doc
    }

    It 'throws in service when array length mismatches Text' {
        $path = Join-Path $TestDrive 'mismatchService.docx'
        $doc = New-OfficeWord -FilePath $path
        { [PSWriteOffice.Services.Word.WordDocumentService]::AddText($doc, $null, @('a','b'), @($true), $null, $null, $null, $null, $null) } | Should -Throw
        Close-OfficeWord -Document $doc
    }

    It 'adds table' {
        $path = Join-Path $TestDrive 'table.docx'
        $doc = New-OfficeWord -FilePath $path
        $data = @([pscustomobject]@{Name='A';Value='1'},[pscustomobject]@{Name='B';Value='2'})
        $table = New-OfficeWordTable -Document $doc -DataTable $data -Suppress
        $table | Should -Not -BeNullOrEmpty
        Close-OfficeWord -Document $doc
    }

    It 'removes header and footer' {
        $path = Join-Path $TestDrive 'hf.docx'
        $doc = New-OfficeWord -FilePath $path
        { Remove-OfficeWordHeader -Document $doc } | Should -Not -Throw
        { Remove-OfficeWordFooter -Document $doc } | Should -Not -Throw
        Close-OfficeWord -Document $doc
    }

    It 'creates list and list item' {
        $path = Join-Path $TestDrive 'list.docx'
        $doc = New-OfficeWord -FilePath $path
        $list = New-OfficeWordList -Document $doc
        New-OfficeWordListItem -List $list -Level 0 -Text 'item'
        $list.Items.Count | Should -BeGreaterThan 0
        Close-OfficeWord -Document $doc
    }

    It 'converts HTML to word document' {
        $path = Join-Path $TestDrive 'html.docx'
        ConvertFrom-HTMLtoWord -OutputFile $path -SourceHTML '<p>Hello</p>'
        $doc = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($path, $false)
        $doc.MainDocumentPart.Document.Body.InnerText | Should -Be 'Hello'
        $doc.Dispose()
    }

    It 'embeds HTML as-is when requested' {
        $path = Join-Path $TestDrive 'htmlasis.docx'
        ConvertFrom-HTMLtoWord -OutputFile $path -SourceHTML '<p>Hello</p>' -Mode AsIs
        $doc = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($path, $false)
        $doc.MainDocumentPart.Document.Body.InnerXml | Should -Match 'altChunk'
        $doc.Dispose()
    }

    It 'throws when saving with null document' {
        { Save-OfficeWord -Document $null } | Should -Throw
    }
}
