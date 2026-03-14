BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Force -Global

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Word reader helpers' {
    It 'finds text and exposes bookmarks/fields' {
        $path = Join-Path $TestDrive 'WordReader.docx'

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph -Text 'Hello world'
        }

        $doc = Get-OfficeWord -Path $path
        try {
            $null = $doc.AddBookmark('Bookmark1')
            $paragraph = $doc.AddParagraph('Page')
            $null = $paragraph.AddField([OfficeIMO.Word.WordFieldType]::Page)
        } finally {
            Close-OfficeWord -Document $doc -Save
        }

        $matches = Find-OfficeWord -Path $path -Text 'Hello'
        $matches.Count | Should -BeGreaterThan 0

        $bookmarks = Get-OfficeWordBookmark -Path $path
        $bookmarks.Name | Should -Contain 'Bookmark1'

        $fields = Get-OfficeWordField -Path $path -FieldType Page
        $fields.Count | Should -BeGreaterThan 0
    }

    It 'reads content controls and table of contents' {
        $path = Join-Path $TestDrive 'WordReaderControls.docx'
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive

        New-OfficeWord -Path $path {
            Add-OfficeWordParagraph { Add-OfficeWordContentControl -Text 'Client' -Alias 'ClientName' -Tag 'ClientTag' }
            Add-OfficeWordParagraph { Add-OfficeWordCheckBox -Checked -Alias 'Approved' -Tag 'ApprovedTag' }
            Add-OfficeWordParagraph { Add-OfficeWordCheckBox -Alias 'Rejected' -Tag 'RejectedTag' }
            Add-OfficeWordParagraph { Add-OfficeWordDatePicker -Date (Get-Date) -Alias 'DueDate' -Tag 'DueTag' }
            Add-OfficeWordParagraph { Add-OfficeWordDropDownList -Items 'Low', 'Medium', 'High' -Alias 'Priority' -Tag 'PriorityTag' }
            Add-OfficeWordParagraph { Add-OfficeWordComboBox -Items 'Red', 'Green' -Alias 'Color' -Tag 'ColorTag' }
            Add-OfficeWordParagraph { Add-OfficeWordPictureControl -Path $imagePath -Alias 'Logo' -Tag 'LogoTag' }
            Add-OfficeWordParagraph { Add-OfficeWordRepeatingSection -SectionTitle 'Items' -Alias 'LineItems' -Tag 'LineItemsTag' }
            $toc = Add-OfficeWordTableOfContent -PassThru
            Set-OfficeWordTableOfContent -TableOfContent $toc -Text 'Contents' -TextNoContent 'No entries'
            Update-OfficeWordTableOfContent -TableOfContent $toc
        }

        $controls = Get-OfficeWordContentControl -Path $path -Alias 'Client*'
        $controls.Count | Should -Be 1

        $checkBoxes = Get-OfficeWordCheckBox -Path $path -Alias 'Approved'
        $checkBoxes.Count | Should -Be 1
        $checkBoxes[0].IsChecked | Should -Be $true

        $unchecked = Get-OfficeWordCheckBox -Path $path -Unchecked
        $unchecked.Count | Should -Be 1

        $datePickers = Get-OfficeWordDatePicker -Path $path -Alias 'DueDate'
        $datePickers.Count | Should -Be 1
        $datePickers[0].Date | Should -Not -BeNullOrEmpty

        $dropDowns = Get-OfficeWordDropDownList -Path $path -Tag 'PriorityTag'
        $dropDowns.Count | Should -Be 1
        $dropDowns[0].Items | Should -Contain 'Medium'

        $comboBoxes = Get-OfficeWordComboBox -Path $path -Alias 'Color'
        $comboBoxes.Count | Should -Be 1
        $comboBoxes[0].Items | Should -Contain 'Red'

        $pictures = Get-OfficeWordPictureControl -Path $path -Alias 'Logo'
        $pictures.Count | Should -Be 1

        $repeating = Get-OfficeWordRepeatingSection -Path $path -Alias 'LineItems'
        $repeating.Count | Should -Be 1

        $toc = Get-OfficeWordTableOfContent -Path $path
        $toc | Should -Not -BeNullOrEmpty
        $toc.Text | Should -Be 'Contents'
        $toc.TextNoContent | Should -Be 'No entries'
    }
}
