Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot\HtmlDocument.docx"
$html = '<h1>Hello</h1><p>Generated from HTML</p>'
ConvertFrom-HTMLtoWord -OutputFile $path -SourceHTML $html
