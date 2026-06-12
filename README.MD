# PSWriteOffice

PowerShell-first document automation for Word, Excel, PowerPoint, PDF, Markdown, CSV, Reader, and Visio, built on top of `OfficeIMO.*`.

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
| PDF | Advanced | New PDFs, text, headings, images, tables, panels, bookmarks, metadata, attachments, stamps, watermarks, forms, merge/split/copy/move pages, HTML conversion, text/image/attachment extraction, compliance inspection |
| Markdown | Solid | Read Markdown, inspect full object trees/headings/front matter/tables, build Markdown with a DSL, render HTML, convert HTML |
| CSV | Solid | Read CSV, emit CSV, object-focused data workflows |
| Reader | New | Unified chunk/document readback over Word, Excel, PowerPoint, Markdown, PDF, and text-like files through `OfficeIMO.Reader` |
| Visio | New | DSL diagram creation, built-in and package-backed stencils, create/load/save `.vsdx`, inspect diagrams, export SVG and PNG through `OfficeIMO.Visio` |

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

```powershell
$chartData = @(
    [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
    [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
)

New-OfficeWord -Path .\Charts.docx {
    Add-OfficeWordChart -Type Pie -Data $chartData -CategoryProperty Region -SeriesProperty Revenue -Title 'Revenue Mix'
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

### PDF

```powershell
$rows = @(
    [PSCustomObject]@{ Area = 'Word'; Status = 'Ready'; Owner = 'Docs' }
    [PSCustomObject]@{ Area = 'PDF'; Status = 'Review'; Owner = 'Adapters' }
)

New-OfficePdf -Path .\Status.pdf {
    Add-OfficePdfHeading -Text 'Documentation Status' -Level 1
    Add-OfficePdfParagraph -Text 'Generated with PSWriteOffice through OfficeIMO.Pdf.'
    Add-OfficePdfTable -InputObject $rows -Property Area,Status,Owner -Header 'Area','Status','Owner' -Align Center
    Add-OfficePdfBookmark -Name 'Status table'
    Set-OfficePdfMetadata -Title 'Documentation Status' -Author 'PSWriteOffice'
}
```

```powershell
Join-OfficePdf -Path .\Cover.pdf, .\Status.pdf -OutputPath .\Report.pdf
Split-OfficePdf -Path .\Report.pdf -OutputDirectory .\Pages
Get-OfficePdfText -Path .\Report.pdf
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

### Reader and Visio

```powershell
Get-OfficeDocumentCapability
Get-OfficeDocumentChunk -Path .\Report.docx
Get-OfficeDocument -Path .\Report.pdf -AsJson -Indented
```

```powershell
New-OfficeVisio -Path .\Diagram.vsdx -Title 'Service map' -RequestRecalcOnOpen {
    VisioRectangle -Key web -Text 'Web' -X 1.5 -Y 4 -Width 1.5 -Height 0.8 -FillColor LightBlue -LineColor SteelBlue
    VisioDiamond -Key decision -Text 'Ready?' -X 4 -Y 4 -Width 1.2 -Height 1 -FillColor '#FFF2CC' -LineColor '#B45309'
    VisioRectangle -Key api -Text 'API' -X 6.2 -Y 4 -Width 1.5 -Height 0.8 -FillColor LightGreen -LineColor SeaGreen
    VisioConnector -From web -To decision -Kind RightAngle -EndArrow Triangle -Label 'check'
    VisioConnector -From decision -To api -Kind RightAngle -EndArrow Triangle -Label 'ship'
}

Get-OfficeVisioInfo -Path .\Diagram.vsdx -AsText
ConvertTo-OfficeVisioSvg -Path .\Diagram.vsdx -OutputPath .\Diagram.svg
ConvertTo-OfficeVisioPng -Path .\Diagram.vsdx -OutputPath .\Diagram.png
```

```powershell
$flow = Get-OfficeVisioStencilCatalog -BuiltIn Flowchart
Find-OfficeVisioStencil -Catalog $flow -Query process -First 5

New-OfficeVisio -Path .\Flow.vsdx -Title 'Stencil flow' -UseMastersByDefault -RequestRecalcOnOpen {
    Import-OfficeVisioStencil -BuiltIn Flowchart -Name Flow -Default
    VisioStencil -Catalog Flow -Stencil process -Key intake -Text 'Intake' -X 1.5 -Y 4 -FillColor '#E0F2FE' -LineColor '#0369A1'
    VisioStencil -Catalog Flow -Stencil decision -Key review -Text 'Review?' -X 4 -Y 4 -FillColor '#FEF3C7' -LineColor '#B45309'
    VisioStencil -Catalog Flow -Stencil data -Key archive -Text 'Archive' -X 6.5 -Y 4 -FillColor '#DCFCE7' -LineColor '#15803D'
    VisioConnector -From intake -To review -Kind RightAngle -EndArrow Triangle -Label 'submit'
    VisioConnector -From review -To archive -Kind RightAngle -EndArrow Triangle -Label 'store'
}
```

```powershell
$catalog = Get-OfficeVisioStencilCatalog -Path .\MyShapes.vssx -CatalogName 'Custom Shapes' -IncludeUnsupportedMasters
Find-OfficeVisioStencil -Catalog $catalog -Query server -First 10

New-OfficeVisio -Path .\CustomStencil.vsdx -UseMastersByDefault {
    Import-OfficeVisioStencil -Catalog $catalog -Name Custom -Default
    VisioStencil -Stencil server -Key server -Text 'Server' -X 2 -Y 4
    VisioStencil -Stencil database -Key database -Text 'Database' -X 5 -Y 4
    VisioConnector -From server -To database -Kind RightAngle -EndArrow Triangle
}
```

More visual Visio examples live in `Examples\Visio`. Run `.\Examples\Visio\Build-VisioShowcase.ps1` to generate editable `.vsdx` files plus SVG/PNG previews for stencil flowcharts, architecture maps, network topology, and package-backed stencil loading.

## Read, Inspect, and Convert 🔍

`PSWriteOffice` is not only about writing files. The module now has stronger read-back and bridge workflows too.

### Excel import/export fast path

```powershell
$data | Export-OfficeExcel -Path .\Report.xlsx -WorksheetName 'Data' -TableName 'Data' -AutoFit -FreezeTopRow
Import-OfficeExcel -Path .\Report.xlsx -WorksheetName 'Data'
```

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
- PDF now has first-class creation, composition, merge/split, page movement, metadata, forms, stamping, attachment, extraction, HTML conversion, and compliance inspection cmdlets. Deeper PDF engine features still belong in `OfficeIMO.Pdf` first.
- `OfficeIMO.Reader` is surfaced through chunk, document-envelope, JSON, and capability-discovery cmdlets.
- `OfficeIMO.Visio` is surfaced through a first diagram DSL plus built-in and external stencil catalogs, create/load/save, deterministic inspection, and SVG/PNG export cmdlets.
- Excel now has `Add-OfficeExcelTableOfContents`, `Get-OfficeExcelRange`, and `Get-OfficeExcelUsedRange`.
- PowerPoint now has section cmdlets, deck-wide text replacement, slide import helpers, slide copy, transitions, slide sizing, theme inspection, theme updates, and layout switching.
- Excel charts can now be finished with `Set-OfficeExcelChartLegend`, `Set-OfficeExcelChartDataLabels`, and `Set-OfficeExcelChartStyle`.
- Excel now has discoverable URL-image insertion plus smarter external hyperlink helpers.
- Excel summary sheets can now auto-link ranges and header-based columns to workbook tabs or external URLs.
- Backlink placement in Excel TOC flows is safer by default, avoiding overwriting active worksheet data.

## Documentation 📚

Use the examples in this README for the first pass, then move to the generated command reference when you need exact parameters and pipeline behavior.

- [OVERVIEW.md](OVERVIEW.md) explains the project scope and design direction.
- [Docs/Readme.md](Docs/Readme.md) is the generated cmdlet index.
- [Docs](Docs) contains one generated page per command.
- [Examples](Examples) contains runnable scripts grouped by Word, Excel, PowerPoint, PDF, Markdown, CSV, and shared samples.
- [Roadmap/PSWriteOffice-CmdletExamplesAndPolishPlan.md](Roadmap/PSWriteOffice-CmdletExamplesAndPolishPlan.md) tracks the documentation quality bar and the next wrapper-polish priorities.

The generated command pages are built from XML comments in the cmdlet source. Prefer contextual examples there too: small workflow snippets with real input data, generated artifacts, readback/validation where useful, and no fake PowerShell compatibility aliases.

## Build and Test 🧪

```powershell
dotnet build .\Sources\PSWriteOffice.sln -c Debug
pwsh -NoLogo -NoProfile -File .\PSWriteOffice.Tests.ps1
pwsh -NoLogo -NoProfile -File .\Build\Validate-PackagedArtefact.ps1
```

Regenerate the module manifest, generated help, packed artifacts, and website API payloads through the build-module lane:

```powershell
pwsh -NoLogo -NoProfile -File .\Build\Manage-PSWriteOffice.ps1
pwsh -NoLogo -NoProfile -File .\Build\Export-ApiDocsSnapshot.ps1
pwsh -NoLogo -NoProfile -File .\Build\Measure-CmdletXmlExamples.ps1 -RequireExamples
```

Development loading is handled through [PSWriteOffice.psm1](PSWriteOffice.psm1), which prefers the local debug build in `Sources\PSWriteOffice\bin\Debug\`.

## License

MIT
