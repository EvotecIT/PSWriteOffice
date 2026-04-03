# PSWriteOffice

PowerShell-first document automation for Word, Excel, PowerPoint, Markdown, and CSV, built on top of `OfficeIMO.*`.

📦 PowerShell Gallery

[![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/PSWriteOffice.svg)](https://www.powershellgallery.com/packages/PSWriteOffice)
[![PowerShell Gallery Preview](https://img.shields.io/powershellgallery/vpre/PSWriteOffice.svg?label=preview)](https://www.powershellgallery.com/packages/PSWriteOffice)
[![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/PSWriteOffice.svg)](https://www.powershellgallery.com/packages/PSWriteOffice)
[![PowerShell Gallery Platforms](https://img.shields.io/powershellgallery/p/PSWriteOffice.svg)](https://www.powershellgallery.com/packages/PSWriteOffice)

🛠️ Project

[![Test PowerShell](https://github.com/EvotecIT/PSWriteOffice/actions/workflows/test-powershell.yml/badge.svg?branch=main)](https://github.com/EvotecIT/PSWriteOffice/actions/workflows/test-powershell.yml)
[![Test .NET Libraries](https://github.com/EvotecIT/PSWriteOffice/actions/workflows/test-dotnet.yml/badge.svg?branch=main)](https://github.com/EvotecIT/PSWriteOffice/actions/workflows/test-dotnet.yml)
[![License](https://img.shields.io/github/license/EvotecIT/PSWriteOffice.svg)](https://github.com/EvotecIT/PSWriteOffice)

## Why PSWriteOffice ✨

`PSWriteOffice` is the PowerShell layer for `OfficeIMO.*`.

- `OfficeIMO` owns the document engine and low-level file-format behavior.
- `PSWriteOffice` owns the PowerShell cmdlets, DSL aliases, packaging, examples, and scripting ergonomics.
- The goal is simple: make Office document automation feel native in PowerShell without requiring Microsoft Office on the machine.

## Platform Areas 🧭

| Area | Status | What it covers now |
| --- | --- | --- |
| Word | Mature | Document creation, readers, bookmarks, content controls, fields, HTML conversion, Markdown conversion |
| Excel | Advanced | Sheets, tables, named ranges, formulas, validation, charts, pivots, comments, TOC/navigation, explicit range readers, chart formatting, smarter links, URL images, summary-sheet linking |
| PowerPoint | Experimental but useful | Slides, titles, text boxes, bullets, notes, layouts, placeholders, sections, text replacement, slide import, slide copy, transitions, sizing |
| Markdown | Solid | Read Markdown, build Markdown with a DSL, render HTML |
| CSV | Solid | Read CSV, emit CSV, object-focused data workflows |

## Quick Start 🚀

### Word

```powershell
New-OfficeWord -Path .\Report.docx {
    WordSection {
        WordHeader { WordParagraph -Text 'Quarterly Report' -Style Heading2 }
        WordFooter { WordPageNumber }
        WordParagraph -Text 'Hello from PSWriteOffice.'
        WordList -Style Bulleted {
            WordListItem -Text 'Alpha'
            WordListItem -Text 'Beta'
        }
    }
}
```

### Excel

```powershell
$data = @(
    [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
    [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
)

New-OfficeExcel -Path .\Report.xlsx {
    ExcelSheet 'Data' {
        ExcelTable -Data $data -TableName 'Sales' -AutoFit
        ExcelNamedRange -Name 'SalesData' -Range 'A1:B3'
    }
    ExcelSheet 'Notes' {
        ExcelRow -Row 1 -Values 'Label', 'Value'
        ExcelRow -Row 2 -Values 'Generated', (Get-Date -Format 'yyyy-MM-dd')
    }
    ExcelTableOfContents -IncludeNamedRanges
}
```

### PowerPoint

```powershell
New-OfficePowerPoint -Path .\Deck.pptx {
    PptSlide {
        PptTitle -Title 'Status Update'
        PptTextBox -Text 'Generated with PSWriteOffice' -X 80 -Y 150 -Width 360 -Height 60
        PptBullets -Bullets 'Wins','Risks','Next Steps' -X 430 -Y 150 -Width 260 -Height 200
        PptNotes -Text 'Keep this under five minutes.'
    }
}
```

```powershell
$ppt = Get-OfficePowerPoint -FilePath .\Deck.pptx
Add-OfficePowerPointSection -Presentation $ppt -Name 'Intro' -StartSlideIndex 0
Rename-OfficePowerPointSection -Presentation $ppt -Name 'Intro' -NewName 'Opening'
Update-OfficePowerPointText -Presentation $ppt -OldValue 'FY24' -NewValue 'FY25' -IncludeNotes
Copy-OfficePowerPointSlide -Presentation $ppt -Index 0
Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideTransition -Transition Fade
Set-OfficePowerPointSlideSize -Presentation $ppt -Preset Screen16x9
Import-OfficePowerPointSlide -Presentation $ppt -SourcePath .\SourceDeck.pptx -SourceIndex 0
```

### PowerPoint theme and layout helpers

```powershell
$ppt = Get-OfficePowerPoint -FilePath .\Deck.pptx
Set-OfficePowerPointThemeColor -Presentation $ppt -Colors @{ Accent1 = '#C00000'; Accent2 = '#00B0F0' } -AllMasters
Set-OfficePowerPointThemeFonts -Presentation $ppt -MajorLatin 'Aptos' -MinorLatin 'Calibri' -AllMasters
Set-OfficePowerPointThemeName -Presentation $ppt -Name 'Contoso Theme' -AllMasters
Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideLayout -LayoutName 'Title and Content'
Get-OfficePowerPointTheme -Presentation $ppt
```

### Excel chart finishing

```powershell
$chart = Add-OfficeExcelChart -TableName 'Sales' -Row 6 -Column 1 -Type Pie -Title 'Revenue Mix' -PassThru
$chart |
    Set-OfficeExcelChartLegend -Position Right |
    Set-OfficeExcelChartDataLabels -ShowValue $true -ShowPercent $true -Position OutsideEnd -NumberFormat '0.0%' -SourceLinked:$false |
    Set-OfficeExcelChartStyle -StyleId 251 -ColorStyleId 10
```

### Excel links and media

```powershell
ExcelSheet 'Data' {
    Set-OfficeExcelSmartHyperlink -Address 'A2' -Url 'https://datatracker.ietf.org/doc/html/rfc7208'
    Set-OfficeExcelHostHyperlink -Address 'B2' -Url 'https://learn.microsoft.com/office/open-xml/'
    Add-OfficeExcelImageFromUrl -Address 'D2' -Url 'https://example.org/logo.png' -WidthPixels 48 -HeightPixels 48
}
```

### Excel internal navigation

```powershell
ExcelSheet 'Summary' {
    Set-OfficeExcelInternalLinks -Range 'D2:D10'
    Set-OfficeExcelInternalLinksByHeader -Header 'Sheet' -TableName 'SummaryTable' -DisplayScript { param($text) "Open $text" }
}
```

### Excel external link automation

```powershell
ExcelSheet 'Summary' {
    Set-OfficeExcelUrlLinksByHeader -Header 'RFC' -TableName 'LinksTable' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" } -TitleScript { param($text) "Open $text" }
    Set-OfficeExcelUrlLinks -Range 'D2:D10' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" }
}
```

### Markdown and CSV

```powershell
New-OfficeMarkdown -Path .\README.md {
    MarkdownHeading -Level 1 -Text 'Report'
    MarkdownParagraph -Text 'Generated by PSWriteOffice.'
}

$data | ConvertTo-OfficeCsv -OutputPath .\export.csv
```

## Read, Inspect, and Convert 🔍

`PSWriteOffice` is not only about writing files. The module now has stronger read-back and bridge workflows too.

### Excel readers

```powershell
Get-OfficeExcelData -Path .\Report.xlsx -Sheet 'Data'
Get-OfficeExcelRange -Path .\Report.xlsx -Sheet 'Data' -Range 'A1:B10'
Get-OfficeExcelUsedRange -Path .\Report.xlsx -Sheet 'Data' -AsDataTable
Get-OfficeExcelNamedRange -Path .\Report.xlsx
Get-OfficeExcelPivotTable -Path .\Report.xlsx
```

### Word bridges

```powershell
$markdown = ConvertTo-OfficeWordMarkdown -Path .\Report.docx
ConvertFrom-OfficeWordMarkdown -Markdown $markdown -OutputPath .\Report-Roundtrip.docx
$doc = Get-OfficeWord -Path .\Report.docx
Update-OfficeWordText -Document $doc -OldValue 'FY24' -NewValue 'FY25'
Close-OfficeWord -Document $doc -Save
```

### Word charts with the current API

```powershell
$doc = New-OfficeWord -Path .\Report.docx
$chart = $doc.AddChart('Revenue Mix')
$chart.AddPie('North America', 125000).
    AddPie('EMEA', 98000).
    AddPie('APAC', 143000) | Out-Null
Close-OfficeWord -Document $doc -Save
```

### Word tables with extra columns

```powershell
$tableData = $data | Select-Object Region, Revenue,
@{ Name = 'RevenueBand'; Expression = { if ($_.Revenue -gt 100000) { 'High' } else { 'Standard' } } }

New-OfficeWord -Path .\Report.docx {
    Add-OfficeWordTable -InputObject $tableData -Style 'GridTable1LightAccent1'
}
```

### PowerPoint inspection

```powershell
$ppt = Get-OfficePowerPoint -FilePath .\Deck.pptx
Get-OfficePowerPointSlide -Presentation $ppt
Get-OfficePowerPointSlideSummary -Presentation $ppt
Get-OfficePowerPointNotes -Presentation $ppt
Get-OfficePowerPointShape -Presentation $ppt -Index 0
```

## Recent Highlights 🆕

- Word to Markdown and Markdown to Word are now surfaced directly through `OfficeIMO.Word.Markdown`.
- Excel now has `Add-OfficeExcelTableOfContents`, `Get-OfficeExcelRange`, and `Get-OfficeExcelUsedRange`.
- PowerPoint now has section cmdlets, deck-wide text replacement, slide import helpers, slide copy, transitions, slide sizing, theme inspection, theme updates, and layout switching.
- Excel charts can now be finished with `Set-OfficeExcelChartLegend`, `Set-OfficeExcelChartDataLabels`, and `Set-OfficeExcelChartStyle`.
- Excel now has discoverable URL-image insertion plus smarter external hyperlink helpers.
- Excel summary sheets can now auto-link ranges and header-based columns to workbook tabs or external URLs.
- Backlink placement in Excel TOC flows is safer by default, avoiding overwriting active worksheet data.

## Documentation 📚

Start here:

- [OVERVIEW.md](OVERVIEW.md)
- [Docs/Readme.md](Docs/Readme.md)

Recommended docs:

- [Docs/New-OfficeWord.md](Docs/New-OfficeWord.md)
- [Docs/Update-OfficeWordText.md](Docs/Update-OfficeWordText.md)
- [Docs/New-OfficeExcel.md](Docs/New-OfficeExcel.md)
- [Docs/New-OfficePowerPoint.md](Docs/New-OfficePowerPoint.md)
- [Docs/ConvertTo-OfficeWordMarkdown.md](Docs/ConvertTo-OfficeWordMarkdown.md)
- [Docs/ConvertFrom-OfficeWordMarkdown.md](Docs/ConvertFrom-OfficeWordMarkdown.md)
- [Docs/Add-OfficeExcelTableOfContents.md](Docs/Add-OfficeExcelTableOfContents.md)
- [Docs/Get-OfficeExcelRange.md](Docs/Get-OfficeExcelRange.md)
- [Docs/Get-OfficeExcelUsedRange.md](Docs/Get-OfficeExcelUsedRange.md)
- [Docs/Set-OfficeExcelChartLegend.md](Docs/Set-OfficeExcelChartLegend.md)
- [Docs/Set-OfficeExcelChartDataLabels.md](Docs/Set-OfficeExcelChartDataLabels.md)
- [Docs/Set-OfficeExcelChartStyle.md](Docs/Set-OfficeExcelChartStyle.md)
- [Docs/Add-OfficeExcelImageFromUrl.md](Docs/Add-OfficeExcelImageFromUrl.md)
- [Docs/Set-OfficeExcelSmartHyperlink.md](Docs/Set-OfficeExcelSmartHyperlink.md)
- [Docs/Set-OfficeExcelHostHyperlink.md](Docs/Set-OfficeExcelHostHyperlink.md)
- [Docs/Set-OfficeExcelInternalLinks.md](Docs/Set-OfficeExcelInternalLinks.md)
- [Docs/Set-OfficeExcelInternalLinksByHeader.md](Docs/Set-OfficeExcelInternalLinksByHeader.md)
- [Docs/Set-OfficeExcelUrlLinks.md](Docs/Set-OfficeExcelUrlLinks.md)
- [Docs/Set-OfficeExcelUrlLinksByHeader.md](Docs/Set-OfficeExcelUrlLinksByHeader.md)
- [Docs/Get-OfficePowerPointSection.md](Docs/Get-OfficePowerPointSection.md)
- [Docs/Get-OfficePowerPointTheme.md](Docs/Get-OfficePowerPointTheme.md)
- [Docs/Set-OfficePowerPointSlideTransition.md](Docs/Set-OfficePowerPointSlideTransition.md)
- [Docs/Set-OfficePowerPointSlideSize.md](Docs/Set-OfficePowerPointSlideSize.md)
- [Docs/Set-OfficePowerPointSlideLayout.md](Docs/Set-OfficePowerPointSlideLayout.md)
- [Docs/Set-OfficePowerPointThemeColor.md](Docs/Set-OfficePowerPointThemeColor.md)
- [Docs/Set-OfficePowerPointThemeFonts.md](Docs/Set-OfficePowerPointThemeFonts.md)
- [Docs/Set-OfficePowerPointThemeName.md](Docs/Set-OfficePowerPointThemeName.md)
- [Docs/Update-OfficePowerPointText.md](Docs/Update-OfficePowerPointText.md)
- [Docs/Copy-OfficePowerPointSlide.md](Docs/Copy-OfficePowerPointSlide.md)
- [Docs/Import-OfficePowerPointSlide.md](Docs/Import-OfficePowerPointSlide.md)

Recommended examples:

- [Examples/Word/Example-WordBasic.ps1](Examples/Word/Example-WordBasic.ps1)
- [Examples/Word/Example-WordReplaceText.ps1](Examples/Word/Example-WordReplaceText.ps1)
- [Examples/Word/Example-WordCharts.ps1](Examples/Word/Example-WordCharts.ps1)
- [Examples/Word/Example-WordLineBreaks.ps1](Examples/Word/Example-WordLineBreaks.ps1)
- [Examples/Word/Example-WordTableCalculatedColumns.ps1](Examples/Word/Example-WordTableCalculatedColumns.ps1)
- [Examples/Word/Example-WordMarkdownConvert.ps1](Examples/Word/Example-WordMarkdownConvert.ps1)
- [Examples/Excel/Example-ExcelBasic.ps1](Examples/Excel/Example-ExcelBasic.ps1)
- [Examples/Excel/Example-ExcelAdvanced.ps1](Examples/Excel/Example-ExcelAdvanced.ps1)
- [Examples/Excel/Example-ExcelNavigationAndRanges.ps1](Examples/Excel/Example-ExcelNavigationAndRanges.ps1)
- [Examples/Excel/Example-ExcelChartFormatting.ps1](Examples/Excel/Example-ExcelChartFormatting.ps1)
- [Examples/Excel/Example-ExcelLinksAndImages.ps1](Examples/Excel/Example-ExcelLinksAndImages.ps1)
- [Examples/Excel/Example-ExcelInternalLinks.ps1](Examples/Excel/Example-ExcelInternalLinks.ps1)
- [Examples/Excel/Example-ExcelUrlLinks.ps1](Examples/Excel/Example-ExcelUrlLinks.ps1)
- [Examples/Markdown/Example-MarkdownDsl.ps1](Examples/Markdown/Example-MarkdownDsl.ps1)
- [Examples/ExamplePowerPoint10-Dsl.ps1](Examples/ExamplePowerPoint10-Dsl.ps1)
- [Examples/ExamplePowerPoint12-LayoutDsl.ps1](Examples/ExamplePowerPoint12-LayoutDsl.ps1)
- [Examples/PowerPoint/Example-PowerPointTransitionsAndSizing.ps1](Examples/PowerPoint/Example-PowerPointTransitionsAndSizing.ps1)
- [Examples/PowerPoint/Example-PowerPointSectionsAndImport.ps1](Examples/PowerPoint/Example-PowerPointSectionsAndImport.ps1)
- [Examples/PowerPoint/Example-PowerPointCopySlides.ps1](Examples/PowerPoint/Example-PowerPointCopySlides.ps1)
- [Examples/PowerPoint/Example-PowerPointThemeAndLayout.ps1](Examples/PowerPoint/Example-PowerPointThemeAndLayout.ps1)

## Build and Test 🧪

```powershell
dotnet build .\Sources\PSWriteOffice.sln -c Debug
pwsh -NoLogo -NoProfile -File .\PSWriteOffice.Tests.ps1
pwsh -NoLogo -NoProfile -File .\Build\Validate-PackagedArtefact.ps1
```

Development loading is handled through [PSWriteOffice.psm1](PSWriteOffice.psm1), which prefers the local debug build in `Sources\PSWriteOffice\bin\Debug\`.

## Roadmap and Next High-Value Additions 🛣️

Near-term focus stays on Word, Excel, PowerPoint, Markdown, and CSV. PDF and Visio remain intentionally out of scope for now because the risk/rework ratio is much higher.

Highest-value next additions from `OfficeIMO` are:

- Higher-level PowerPoint background/design helpers that build on the newer theme, layout, and slide reuse workflows
- Additional Word and Excel inspection helpers where the underlying C# APIs are already stable
- More Excel dashboard/reporting helpers on top of the newer link, URL, and TOC primitives

The general direction is clear: keep PSWriteOffice thin, predictable, and PowerShell-native while steadily exposing the strongest `OfficeIMO` capabilities.

## Dependency Model 🔗

`PSWriteOffice` should stay aligned with `OfficeIMO.*` rather than re-owning lower-level behavior or duplicating file-format logic.

- `OfficeIMO.Word`
- `OfficeIMO.Word.Markdown`
- `OfficeIMO.Excel`
- `OfficeIMO.PowerPoint`
- `OfficeIMO.Markdown`
- `OfficeIMO.CSV`
- `OfficeIMO.Word.Html`

That keeps feature ownership centralized and reduces drift between the PowerShell layer and the underlying document engine.

## License

MIT
