Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Word-Find.docx'

New-OfficeWord -Path $path {
    Add-OfficeWordParagraph -Text 'Hello from PSWriteOffice'
} | Out-Null

$doc = Get-OfficeWord -Path $path
try {
    $null = $doc.AddBookmark('Bookmark1')
    $paragraph = $doc.AddParagraph('Page')
    $null = $paragraph.AddField([OfficeIMO.Word.WordFieldType]::Page)
} finally {
    Close-OfficeWord -Document $doc -Save
}

Find-OfficeWord -Path $path -Text 'Hello'
Get-OfficeWordBookmark -Path $path
Get-OfficeWordField -Path $path -FieldType Page
