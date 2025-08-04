Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$pathAsIs = "$PSScriptRoot\Documents\HtmlAsIs.docx"
ConvertFrom-HTMLtoWord -OutputFile $pathAsIs -SourceHTML '<p>Hello <b>World</b></p>' -Mode AsIs -Show

$pathParsed = "$PSScriptRoot\Documents\HtmlParsed.docx"
ConvertFrom-HTMLtoWord -OutputFile $pathParsed -SourceHTML '<p>Hello <b>World</b></p>' -Mode Parse -Show
