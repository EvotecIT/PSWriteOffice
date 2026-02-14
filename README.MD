# PSWriteOffice - PowerShell Module

<p align="center">
  <a href="https://dev.azure.com/evotecpl/PSWriteOffice/_build/results?buildId=latest"><img src="https://img.shields.io/azure-devops/build/evotecpl/39c74615-8f34-4af0-a835-68dc33f9214f/14?label=Azure%20Pipelines&style=flat-square"></a>
  <a href="https://www.powershellgallery.com/packages/PSWriteOffice"><img src="https://img.shields.io/powershellgallery/v/PSWriteOffice.svg?style=flat-square"></a>
  <a href="https://www.powershellgallery.com/packages/PSWriteOffice"><img src="https://img.shields.io/powershellgallery/v/PSWriteOffice.svg?label=powershell%20gallery%20preview&colorB=yellow&style=flat-square&include_prereleases"></a>
  <a href="https://github.com/EvotecIT/PSWriteOffice"><img src="https://img.shields.io/github/license/EvotecIT/PSWriteOffice.svg?style=flat-square"></a>
</p>

<p align="center">
  <a href="https://www.powershellgallery.com/packages/PSWriteOffice"><img src="https://img.shields.io/powershellgallery/p/PSWriteOffice.svg?style=flat-square"></a>
  <a href="https://github.com/EvotecIT/PSWriteOffice"><img src="https://img.shields.io/github/languages/top/evotecit/PSWriteOffice.svg?style=flat-square"></a>
  <a href="https://github.com/EvotecIT/PSWriteOffice"><img src="https://img.shields.io/github/languages/code-size/evotecit/PSWriteOffice.svg?style=flat-square"></a>
  <a href="https://www.powershellgallery.com/packages/PSWriteOffice"><img src="https://img.shields.io/powershellgallery/dt/PSWriteOffice.svg?style=flat-square"></a>
</p>

<p align="center">
  <a href="https://twitter.com/PrzemyslawKlys"><img src="https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=flat-square&logo=twitter"></a>
  <a href="https://evotec.xyz/hub"><img src="https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg?style=flat-square"></a>
  <a href="https://www.linkedin.com/in/pklys"><img src="https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn&style=flat-square"></a>
</p>

`PSWriteOffice` is a cross-platform PowerShell module for creating and reading Microsoft Office documents (`Word`, `Excel`, `PowerPoint` - experimental), plus `Markdown` and `CSV`, powered by OfficeIMO.*.
It provides verb-noun cmdlets and a scriptblock-first DSL while OfficeIMO owns the core engine and tests.

This module is built using following libraries:
- [Open XML SDK for .NET](https://github.com/OfficeDev/Open-XML-SDK) licensed MIT - Microsoft's library doing heavy lifting for all other Open XML projects.
- [OfficeIMO.Word.Html](https://github.com/EvotecIT/OfficeIMO) - OfficeIMO HTML converters for Word.
- [OfficeIMO](https://github.com/EvotecIT/OfficeIMO) - core engines: Word, Excel, PowerPoint, Markdown, CSV.
- [SixLabors.ImageSharp](https://github.com/SixLabors/ImageSharp) licensed Apache 2.0 - library for image processing.

HTML converters (OfficeIMO.Word.Html):
- `ConvertTo-OfficeWordHtml -Path .\report.docx -OutputPath .\report.html`
- `ConvertFrom-OfficeWordHtml -Path .\snippet.html -OutputPath .\report.docx`

CSV helpers:
- `Get-OfficeCsv -Path .\data.csv`
- `Get-OfficeCsvData -Path .\data.csv`
- `ConvertTo-OfficeCsv -OutputPath .\export.csv`

Markdown helpers:
- `Get-OfficeMarkdown -Path .\README.md`
- `ConvertTo-OfficeMarkdown` (from objects)
- `ConvertTo-OfficeMarkdownHtml -Text $md -DocumentMode`

Word DSL quick start:
- `New-OfficeWord -Path .\Report.docx { WordSection { WordParagraph 'Hello' } }`
- `Get-OfficeWordParagraph -Path .\Report.docx | Select-Object -First 3`

Excel DSL quick start:
- `New-OfficeExcel -Path .\Report.xlsx { ExcelSheet 'Data' { Set-OfficeExcelRow -Row 1 -Values 'Name','Value' } }`
- `Get-OfficeExcelNamedRange -Path .\Report.xlsx`
