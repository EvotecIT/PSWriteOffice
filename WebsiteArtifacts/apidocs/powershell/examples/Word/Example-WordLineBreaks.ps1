Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Example-WordLineBreaks.docx'
$document = New-OfficeWord -Path $path

# AddBreak() creates a same-paragraph line break similar to Shift+Enter in Word.
$paragraph = $document.AddParagraph('Line 1 in the same paragraph')
$null = $paragraph.AddBreak()
$null = $paragraph.AddText('Line 2 after AddBreak()')
$null = $paragraph.AddBreak()
$null = $paragraph.AddText('Line 3 still in the same paragraph')

# AddParagraph() creates a new paragraph, so an empty paragraph gives a visible blank line.
$null = $document.AddParagraph()
$null = $document.AddParagraph('This text comes after an empty paragraph break.')

Close-OfficeWord -Document $document -Save

Write-Host "Document saved to $path"
