# PSWriteOffice Overview

Use this page as the curated starting point for the repo. The generated command reference lives in [Docs/Readme.md](Docs/Readme.md), while this overview focuses on what to use first, where the module is strongest, and which examples match the tested surface.

## Start Here 🚀

If you want to create documents:

- Word: [Docs/New-OfficeWord.md](Docs/New-OfficeWord.md)
- Excel: [Docs/New-OfficeExcel.md](Docs/New-OfficeExcel.md)
- PowerPoint: [Docs/New-OfficePowerPoint.md](Docs/New-OfficePowerPoint.md)
- Markdown: [Docs/New-OfficeMarkdown.md](Docs/New-OfficeMarkdown.md)
- CSV: [Docs/ConvertTo-OfficeCsv.md](Docs/ConvertTo-OfficeCsv.md)

If you want to inspect existing files:

- Word: [Docs/Get-OfficeWordParagraph.md](Docs/Get-OfficeWordParagraph.md), [Docs/Find-OfficeWord.md](Docs/Find-OfficeWord.md)
- Excel: [Docs/Get-OfficeExcelData.md](Docs/Get-OfficeExcelData.md), [Docs/Get-OfficeExcelRange.md](Docs/Get-OfficeExcelRange.md), [Docs/Get-OfficeExcelUsedRange.md](Docs/Get-OfficeExcelUsedRange.md), [Docs/Get-OfficeExcelNamedRange.md](Docs/Get-OfficeExcelNamedRange.md)
- PowerPoint: [Docs/Get-OfficePowerPoint.md](Docs/Get-OfficePowerPoint.md), [Docs/Get-OfficePowerPointSlideSummary.md](Docs/Get-OfficePowerPointSlideSummary.md), [Docs/Get-OfficePowerPointShape.md](Docs/Get-OfficePowerPointShape.md), [Docs/Get-OfficePowerPointSection.md](Docs/Get-OfficePowerPointSection.md), [Docs/Get-OfficePowerPointTheme.md](Docs/Get-OfficePowerPointTheme.md)
- Markdown: [Docs/Get-OfficeMarkdown.md](Docs/Get-OfficeMarkdown.md)
- CSV: [Docs/Get-OfficeCsv.md](Docs/Get-OfficeCsv.md), [Docs/Get-OfficeCsvData.md](Docs/Get-OfficeCsvData.md)

If you want conversion and bridge workflows:

- Word to HTML: [Docs/ConvertTo-OfficeWordHtml.md](Docs/ConvertTo-OfficeWordHtml.md)
- HTML to Word: [Docs/ConvertFrom-OfficeWordHtml.md](Docs/ConvertFrom-OfficeWordHtml.md)
- Word to Markdown: [Docs/ConvertTo-OfficeWordMarkdown.md](Docs/ConvertTo-OfficeWordMarkdown.md)
- Markdown to Word: [Docs/ConvertFrom-OfficeWordMarkdown.md](Docs/ConvertFrom-OfficeWordMarkdown.md)
- Word text replacement: [Docs/Update-OfficeWordText.md](Docs/Update-OfficeWordText.md)
- Word chart example: [Examples/Word/Example-WordCharts.ps1](Examples/Word/Example-WordCharts.ps1)
- Word table projection example: [Examples/Word/Example-WordTableCalculatedColumns.ps1](Examples/Word/Example-WordTableCalculatedColumns.ps1)
- Excel navigation sheet: [Docs/Add-OfficeExcelTableOfContents.md](Docs/Add-OfficeExcelTableOfContents.md)
- Excel chart finishing: [Docs/Set-OfficeExcelChartLegend.md](Docs/Set-OfficeExcelChartLegend.md), [Docs/Set-OfficeExcelChartDataLabels.md](Docs/Set-OfficeExcelChartDataLabels.md), [Docs/Set-OfficeExcelChartStyle.md](Docs/Set-OfficeExcelChartStyle.md)
- Excel links and media: [Docs/Add-OfficeExcelImageFromUrl.md](Docs/Add-OfficeExcelImageFromUrl.md), [Docs/Set-OfficeExcelSmartHyperlink.md](Docs/Set-OfficeExcelSmartHyperlink.md), [Docs/Set-OfficeExcelHostHyperlink.md](Docs/Set-OfficeExcelHostHyperlink.md)
- Excel internal navigation: [Docs/Set-OfficeExcelInternalLinks.md](Docs/Set-OfficeExcelInternalLinks.md), [Docs/Set-OfficeExcelInternalLinksByHeader.md](Docs/Set-OfficeExcelInternalLinksByHeader.md)
- Excel external reporting links: [Docs/Set-OfficeExcelUrlLinks.md](Docs/Set-OfficeExcelUrlLinks.md), [Docs/Set-OfficeExcelUrlLinksByHeader.md](Docs/Set-OfficeExcelUrlLinksByHeader.md)
- PowerPoint text replacement: [Docs/Update-OfficePowerPointText.md](Docs/Update-OfficePowerPointText.md)
- PowerPoint transitions and sizing: [Docs/Set-OfficePowerPointSlideTransition.md](Docs/Set-OfficePowerPointSlideTransition.md), [Docs/Set-OfficePowerPointSlideSize.md](Docs/Set-OfficePowerPointSlideSize.md)
- PowerPoint themes and layouts: [Docs/Get-OfficePowerPointTheme.md](Docs/Get-OfficePowerPointTheme.md), [Docs/Set-OfficePowerPointThemeColor.md](Docs/Set-OfficePowerPointThemeColor.md), [Docs/Set-OfficePowerPointThemeFonts.md](Docs/Set-OfficePowerPointThemeFonts.md), [Docs/Set-OfficePowerPointThemeName.md](Docs/Set-OfficePowerPointThemeName.md), [Docs/Set-OfficePowerPointSlideLayout.md](Docs/Set-OfficePowerPointSlideLayout.md)
- PowerPoint slide copy: [Docs/Copy-OfficePowerPointSlide.md](Docs/Copy-OfficePowerPointSlide.md)
- PowerPoint slide import: [Docs/Import-OfficePowerPointSlide.md](Docs/Import-OfficePowerPointSlide.md)

## Module Areas 🧭

| Area | Status | Start here |
| --- | --- | --- |
| Word | Mature | [Docs/New-OfficeWord.md](Docs/New-OfficeWord.md), [Docs/Get-OfficeWord.md](Docs/Get-OfficeWord.md), [Docs/ConvertTo-OfficeWordMarkdown.md](Docs/ConvertTo-OfficeWordMarkdown.md) |
| Excel | Advanced | [Docs/New-OfficeExcel.md](Docs/New-OfficeExcel.md), [Docs/Get-OfficeExcel.md](Docs/Get-OfficeExcel.md), [Docs/Get-OfficeExcelRange.md](Docs/Get-OfficeExcelRange.md), [Docs/Add-OfficeExcelTableOfContents.md](Docs/Add-OfficeExcelTableOfContents.md), [Docs/Set-OfficeExcelChartLegend.md](Docs/Set-OfficeExcelChartLegend.md), [Docs/Set-OfficeExcelSmartHyperlink.md](Docs/Set-OfficeExcelSmartHyperlink.md), [Docs/Set-OfficeExcelInternalLinks.md](Docs/Set-OfficeExcelInternalLinks.md), [Docs/Set-OfficeExcelUrlLinks.md](Docs/Set-OfficeExcelUrlLinks.md) |
| PowerPoint | Improving fast | [Docs/New-OfficePowerPoint.md](Docs/New-OfficePowerPoint.md), [Docs/Get-OfficePowerPointSlideSummary.md](Docs/Get-OfficePowerPointSlideSummary.md), [Docs/Get-OfficePowerPointShape.md](Docs/Get-OfficePowerPointShape.md), [Docs/Get-OfficePowerPointSection.md](Docs/Get-OfficePowerPointSection.md), [Docs/Get-OfficePowerPointTheme.md](Docs/Get-OfficePowerPointTheme.md), [Docs/Copy-OfficePowerPointSlide.md](Docs/Copy-OfficePowerPointSlide.md), [Docs/Set-OfficePowerPointSlideTransition.md](Docs/Set-OfficePowerPointSlideTransition.md), [Docs/Set-OfficePowerPointSlideLayout.md](Docs/Set-OfficePowerPointSlideLayout.md) |
| Markdown | Solid | [Docs/New-OfficeMarkdown.md](Docs/New-OfficeMarkdown.md), [Docs/Get-OfficeMarkdown.md](Docs/Get-OfficeMarkdown.md), [Docs/ConvertTo-OfficeMarkdownHtml.md](Docs/ConvertTo-OfficeMarkdownHtml.md) |
| CSV | Solid | [Docs/Get-OfficeCsv.md](Docs/Get-OfficeCsv.md), [Docs/Get-OfficeCsvData.md](Docs/Get-OfficeCsvData.md), [Docs/ConvertTo-OfficeCsv.md](Docs/ConvertTo-OfficeCsv.md) |

## Recommended Examples 🧪

Word:

- [Examples/Word/Example-WordBasic.ps1](Examples/Word/Example-WordBasic.ps1)
- [Examples/Word/Example-WordAdvanced.ps1](Examples/Word/Example-WordAdvanced.ps1)
- [Examples/Word/Example-WordReplaceText.ps1](Examples/Word/Example-WordReplaceText.ps1)
- [Examples/Word/Example-WordCharts.ps1](Examples/Word/Example-WordCharts.ps1)
- [Examples/Word/Example-WordLineBreaks.ps1](Examples/Word/Example-WordLineBreaks.ps1)
- [Examples/Word/Example-WordTableCalculatedColumns.ps1](Examples/Word/Example-WordTableCalculatedColumns.ps1)
- [Examples/Word/Example-WordMarkdownConvert.ps1](Examples/Word/Example-WordMarkdownConvert.ps1)

Excel:

- [Examples/Excel/Example-ExcelBasic.ps1](Examples/Excel/Example-ExcelBasic.ps1)
- [Examples/Excel/Example-ExcelAdvanced.ps1](Examples/Excel/Example-ExcelAdvanced.ps1)
- [Examples/Excel/Example-ExcelReadObjects.ps1](Examples/Excel/Example-ExcelReadObjects.ps1)
- [Examples/Excel/Example-ExcelNavigationAndRanges.ps1](Examples/Excel/Example-ExcelNavigationAndRanges.ps1)
- [Examples/Excel/Example-ExcelChartFormatting.ps1](Examples/Excel/Example-ExcelChartFormatting.ps1)
- [Examples/Excel/Example-ExcelLinksAndImages.ps1](Examples/Excel/Example-ExcelLinksAndImages.ps1)
- [Examples/Excel/Example-ExcelInternalLinks.ps1](Examples/Excel/Example-ExcelInternalLinks.ps1)
- [Examples/Excel/Example-ExcelUrlLinks.ps1](Examples/Excel/Example-ExcelUrlLinks.ps1)

PowerPoint:

- [Examples/ExamplePowerPoint08-TablesAndShapes.ps1](Examples/ExamplePowerPoint08-TablesAndShapes.ps1)
- [Examples/ExamplePowerPoint10-Dsl.ps1](Examples/ExamplePowerPoint10-Dsl.ps1)
- [Examples/ExamplePowerPoint12-LayoutDsl.ps1](Examples/ExamplePowerPoint12-LayoutDsl.ps1)
- [Examples/PowerPoint/Example-PowerPointTransitionsAndSizing.ps1](Examples/PowerPoint/Example-PowerPointTransitionsAndSizing.ps1)
- [Examples/PowerPoint/Example-PowerPointSectionsAndImport.ps1](Examples/PowerPoint/Example-PowerPointSectionsAndImport.ps1)
- [Examples/PowerPoint/Example-PowerPointCopySlides.ps1](Examples/PowerPoint/Example-PowerPointCopySlides.ps1)
- [Examples/PowerPoint/Example-PowerPointThemeAndLayout.ps1](Examples/PowerPoint/Example-PowerPointThemeAndLayout.ps1)

Markdown and CSV:

- [Examples/Markdown/Example-MarkdownBasic.ps1](Examples/Markdown/Example-MarkdownBasic.ps1)
- [Examples/Markdown/Example-MarkdownAdvanced.ps1](Examples/Markdown/Example-MarkdownAdvanced.ps1)
- [Examples/Csv/Example-CsvBasic.ps1](Examples/Csv/Example-CsvBasic.ps1)
- [Examples/Csv/Example-CsvAdvanced.ps1](Examples/Csv/Example-CsvAdvanced.ps1)

## Current Direction 🧱

- Keep PSWriteOffice as a thin PowerShell wrapper over `OfficeIMO.*`
- Expand the strongest Excel and PowerPoint capabilities already proven in the C# libraries
- Prefer object-first read and inspection helpers over overly clever abstractions
- Keep docs and examples aligned with the actually tested cmdlet surface

## High-Value Next Additions 🛣️

- PowerPoint background/design helpers on top of the new theme/layout surface
- Additional object-first inspection helpers in Excel and Word
- More Excel dashboard/reporting helpers on top of the link, URL, and TOC surface

PDF and Visio are intentionally deferred for now.

## Generated Command Help 📚

The full generated command reference is in [Docs/Readme.md](Docs/Readme.md). Keep long-lived hand-written guidance outside `Docs/Readme.md`, and let the generated help pages continue to cover the full cmdlet inventory.
