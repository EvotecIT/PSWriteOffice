# PSWriteOffice Overview

`PSWriteOffice` is the PowerShell layer for `OfficeIMO.*`.

Use this page as the stable starting point for the repo. The generated command help lives in [Docs/Readme.md](Docs/Readme.md), while this page focuses on what the module does today, where to start, and which examples match the tested surface.

## Install

```powershell
Install-Module PSWriteOffice -Scope CurrentUser
Import-Module PSWriteOffice
```

## Module Areas

| Area | Status | Start here |
| --- | --- | --- |
| Word | Mature | [New-OfficeWord](Docs/New-OfficeWord.md), [Get-OfficeWord](Docs/Get-OfficeWord.md), [ConvertTo-OfficeWordHtml](Docs/ConvertTo-OfficeWordHtml.md) |
| Excel | Advanced | [New-OfficeExcel](Docs/New-OfficeExcel.md), [Get-OfficeExcel](Docs/Get-OfficeExcel.md), [Add-OfficeExcelTable](Docs/Add-OfficeExcelTable.md) |
| PowerPoint | Improving fast | [New-OfficePowerPoint](Docs/New-OfficePowerPoint.md), [Get-OfficePowerPointSlide](Docs/Get-OfficePowerPointSlide.md), [Get-OfficePowerPointShape](Docs/Get-OfficePowerPointShape.md) |
| Markdown | Solid | [New-OfficeMarkdown](Docs/New-OfficeMarkdown.md), [Get-OfficeMarkdown](Docs/Get-OfficeMarkdown.md), [ConvertTo-OfficeMarkdownHtml](Docs/ConvertTo-OfficeMarkdownHtml.md) |
| CSV | Solid | [Get-OfficeCsv](Docs/Get-OfficeCsv.md), [Get-OfficeCsvData](Docs/Get-OfficeCsvData.md), [ConvertTo-OfficeCsv](Docs/ConvertTo-OfficeCsv.md) |

## Suggested Entry Points

If you want to create documents:

- Word: [New-OfficeWord](Docs/New-OfficeWord.md)
- Excel: [New-OfficeExcel](Docs/New-OfficeExcel.md)
- PowerPoint: [New-OfficePowerPoint](Docs/New-OfficePowerPoint.md)
- Markdown: [New-OfficeMarkdown](Docs/New-OfficeMarkdown.md)
- CSV: [ConvertTo-OfficeCsv](Docs/ConvertTo-OfficeCsv.md)

If you want to inspect existing files:

- Word: [Get-OfficeWordParagraph](Docs/Get-OfficeWordParagraph.md), [Find-OfficeWord](Docs/Find-OfficeWord.md)
- Excel: [Get-OfficeExcelData](Docs/Get-OfficeExcelData.md), [Get-OfficeExcelNamedRange](Docs/Get-OfficeExcelNamedRange.md), [Get-OfficeExcelPivotTable](Docs/Get-OfficeExcelPivotTable.md)
- PowerPoint: [Get-OfficePowerPoint](Docs/Get-OfficePowerPoint.md), [Get-OfficePowerPointSlide](Docs/Get-OfficePowerPointSlide.md), [Get-OfficePowerPointPlaceholder](Docs/Get-OfficePowerPointPlaceholder.md), [Get-OfficePowerPointNotes](Docs/Get-OfficePowerPointNotes.md), [Get-OfficePowerPointShape](Docs/Get-OfficePowerPointShape.md)
- Markdown: [Get-OfficeMarkdown](Docs/Get-OfficeMarkdown.md)
- CSV: [Get-OfficeCsv](Docs/Get-OfficeCsv.md), [Get-OfficeCsvData](Docs/Get-OfficeCsvData.md)

## Example Scripts

- Word
  - [Examples/Word/Example-WordBasic.ps1](Examples/Word/Example-WordBasic.ps1)
  - [Examples/Word/Example-WordAdvanced.ps1](Examples/Word/Example-WordAdvanced.ps1)
  - [Examples/Word/Example-WordReadDocument.ps1](Examples/Word/Example-WordReadDocument.ps1)
- Excel
  - [Examples/Excel/Example-ExcelBasic.ps1](Examples/Excel/Example-ExcelBasic.ps1)
  - [Examples/Excel/Example-ExcelAdvanced.ps1](Examples/Excel/Example-ExcelAdvanced.ps1)
  - [Examples/Excel/Example-ExcelReadObjects.ps1](Examples/Excel/Example-ExcelReadObjects.ps1)
- PowerPoint
  - [Examples/ExamplePowerPoint08-TablesAndShapes.ps1](Examples/ExamplePowerPoint08-TablesAndShapes.ps1)
  - [Examples/ExamplePowerPoint10-Dsl.ps1](Examples/ExamplePowerPoint10-Dsl.ps1)
  - [Examples/ExamplePowerPoint12-LayoutDsl.ps1](Examples/ExamplePowerPoint12-LayoutDsl.ps1)
- Markdown
  - [Examples/Markdown/Example-MarkdownBasic.ps1](Examples/Markdown/Example-MarkdownBasic.ps1)
  - [Examples/Markdown/Example-MarkdownAdvanced.ps1](Examples/Markdown/Example-MarkdownAdvanced.ps1)
- CSV
  - [Examples/Csv/Example-CsvBasic.ps1](Examples/Csv/Example-CsvBasic.ps1)
  - [Examples/Csv/Example-CsvAdvanced.ps1](Examples/Csv/Example-CsvAdvanced.ps1)

## Generated Command Help

The full generated command reference is in [Docs/Readme.md](Docs/Readme.md). That folder is builder-generated, so keep long-lived hand-written guidance outside `Docs/` and let the generated help pages cover the full cmdlet list.
