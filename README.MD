# PSWriteOffice

PowerShell-first document automation for Office, PDF, OpenDocument, email, Markdown, AsciiDoc, LaTeX, CSV, Reader, and Visio, built on top of `OfficeIMO.*`.

📦 PowerShell Gallery

[![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/PSWriteOffice.svg)](https://www.powershellgallery.com/packages/PSWriteOffice)
[![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/PSWriteOffice.svg)](https://www.powershellgallery.com/packages/PSWriteOffice)
[![PowerShell Gallery Platforms](https://img.shields.io/powershellgallery/p/PSWriteOffice.svg)](https://www.powershellgallery.com/packages/PSWriteOffice)

🛠️ Project

[![Test PowerShell](https://github.com/EvotecIT/PSWriteOffice/actions/workflows/test-powershell.yml/badge.svg?branch=main)](https://github.com/EvotecIT/PSWriteOffice/actions/workflows/test-powershell.yml)
[![Test .NET Libraries](https://github.com/EvotecIT/PSWriteOffice/actions/workflows/test-dotnet.yml/badge.svg?branch=main)](https://github.com/EvotecIT/PSWriteOffice/actions/workflows/test-dotnet.yml)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

## Why PSWriteOffice ✨

`PSWriteOffice` brings `OfficeIMO.*` document automation to PowerShell through cmdlets, DSL aliases, examples, and scripting-friendly defaults.

The goal is simple: make Office document automation feel native in PowerShell without requiring Microsoft Office on the machine.

Start here when you want to create polished documents from PowerShell: workbooks with tables, charts, links, images, pivots, and navigation; Word reports with sections, headers, footers, tables, charts, and conversions; PowerPoint decks with reusable slides, notes, sections, themes, and transitions; PDFs with composition, extraction, forms, stamps, and metadata; RTF documents and bridges; Markdown and CSV workflows; Visio diagrams; and unified readback for existing files.

## Platform Areas 🧭

| Area | Status | What it covers now |
| --- | --- | --- |
| Word | Mature | Creation and editing, comparison and redlines, comments and revision reports, HTML/Markdown/OpenDocument conversion, Google Docs planning/export, and PNG/SVG rendering |
| Excel | Advanced | Sheets, tables, formulas, validation, charts, pivots, comments, navigation, explicit readers, OpenDocument conversion, Google Sheets planning/export, and PNG/SVG rendering |
| PowerPoint | Advanced | Slides, layouts, themes, notes, sections, transitions, import/copy, accessibility and preflight inspection, OpenDocument conversion, and PNG/SVG rendering |
| PDF | Advanced | Composition, forms, sanitization, XFDF, visual comparison, interaction maps, layout overlays, OCR merge, rewrite proof, extraction, compliance inspection, and PNG/SVG rendering |
| OpenDocument | Native | Create, load, save, and convert ODT, ODS, and ODP with conversion diagnostics and fail-on-loss controls |
| Email | Native | Read and write EML, MSG, and TNEF; read and write mbox mailboxes with bounded diagnostics |
| RTF | Solid | Create/read/update `.rtf`, replace text, append paragraphs, edit document properties, convert RTF to Word/HTML/PDF, convert Word/HTML to RTF, and read semantic chunks through Reader |
| Markdown and text formats | Solid | Native Markdown, AsciiDoc, and LaTeX parsing plus loss-aware conversions between Markdown and AsciiDoc/LaTeX |
| CSV | Solid | Read CSV, emit CSV, object-focused data workflows |
| Reader | Advanced | Content detection, structured extraction, token-aware hierarchy, processors, OCR adapters, and bounded ordered batch reads across Office, OpenDocument, email, PDF, markup, archives, and text formats |
| Visio | New | DSL diagram creation, built-in and package-backed stencils, create/load/save `.vsdx`, inspect diagrams, generate galleries, export stencil preview galleries, and export SVG/PNG through `OfficeIMO.Visio` |
| Confluence Cloud | New | ADF and storage-format page reads, dry-run create/update plans, loss-aware Markdown/HTML conversion, managed sections, and attachment transfer |

## Advanced document workflows

Reader commands return OfficeIMO's typed results, so scripts can use detection evidence, structured records, token-aware chunks, diagnostics, and deterministic batch ordering directly:

```powershell
$detection = Get-OfficeDocumentDetection -Path .\Upload.bin -Mode PreferContent
$structured = Get-OfficeDocumentStructured -Path .\Report.docx
$hierarchy = Get-OfficeDocumentHierarchy -Path .\Handbook.pdf
$document = Get-OfficeDocument -Path .\Policy.docx -IncludePageLocations
$matches = $document | Search-OfficeDocument -Query 'retention period'
$pageMarkdown = $document | Get-OfficeDocumentPageMarkdown -AsString
$reader = New-OfficeDocumentReader -TesseractOptions $tesseract -MaxConcurrentReads 4
Get-ChildItem .\Inbox -File | Get-OfficeDocumentBatch -Reader $reader
```

Reader search and page Markdown use OfficeIMO's normalized blocks, page locations, and provenance. `-IncludePageLocations` opts Word into computed pagination and RTF into explicit-break reconstruction; fixed-layout and logical-container formats expose their native page-like locations without it.

Format-specific Reader settings are captured when handlers are registered. Pass an `OfficeIMO.Reader.All.ReaderAllOptions` instance to `New-OfficeDocumentReader -ReaderAllOptions`; keep ordinary `ReaderOptions` on the individual read command. A caller-supplied reader is treated as immutable, so PSWriteOffice will not silently replace its registered handler settings.

Image export uses the same `OfficeImageExportResult` contract for Word, Excel, PowerPoint, HTML, and PDF output:

```powershell
Export-OfficeWordImage -Path .\Report.docx -OutputPath .\Report.svg -Format Svg
Export-OfficeExcelImage -Path .\Data.xlsx -OutputPath .\WorkbookImages -Format Png
Export-OfficePowerPointImage -Path .\Deck.pptx -OutputPath .\SlideImages -Format Svg
Export-OfficeHtmlImage -Path .\Report.html -OutputPath .\Report.png
Export-OfficePdfImage -Path .\Report.pdf -OutputPath .\PdfPages -PageRange '1-3'
```

Conversions that can simplify or omit source features return typed reports. Use `-FailOnLoss` when approximation is not acceptable:

```powershell
$odf = ConvertTo-OfficeOpenDocument -Path .\Report.docx -OutputPath .\Report.odt -FailOnLoss
$asciiDoc = Get-OfficeMarkdown -Path .\Guide.md |
    ConvertFrom-OfficeAsciiDocMarkdown -OutputPath .\Guide.adoc -FailOnLoss
$latex = Get-OfficeMarkdown -Path .\Paper.md |
    ConvertFrom-OfficeLatexMarkdown -OutputPath .\Paper.tex -FailOnLoss
```

Native email and Google Workspace commands keep provider and credential ownership explicit:

```powershell
$message = Get-OfficeEmail -Path .\Message.msg -AsResult
$mailbox = Get-OfficeEmailMailbox -Path .\Archive.mbox -AsResult
$docsPlan = Export-OfficeWordGoogleDocument -Path .\Report.docx -PlanOnly
$sheetsBatch = Export-OfficeExcelGoogleSpreadsheet -Path .\Data.xlsx -AsBatch
```

Confluence writes can be reviewed before a session or tenant is involved:

```powershell
$plan = Publish-OfficeConfluencePage `
    -SpaceId 42 `
    -Title 'Daily status' `
    -Content "# Ready`n`nGenerated from the current report." `
    -PlanOnly

$plan.Method
$plan.RelativeUri
$plan.Payload
```

For a live write, create a session with an Atlassian email/API-token credential or an OAuth bearer token, then pass it through `-Session`. OAuth uses Atlassian's cloud-ID gateway, so bearer sessions require `-CloudId`:

```powershell
$session = New-OfficeConfluenceSession `
    -SiteUri 'https://example.atlassian.net/' `
    -AccessToken $secureAccessToken `
    -CloudId $cloudId

Publish-OfficeConfluencePage -Session $session -SpaceId 42 -Title 'Daily status' -Content $markdown

# Review a destructive request first; omit -PlanOnly only when ready.
Remove-OfficeConfluencePage -PageId 12345 -Purge -PlanOnly
```

`Set-OfficeConfluenceManagedSection` provides a pure before/after operation for generated sections so surrounding owner-authored storage content can be preserved. File-based attachment uploads and downloads stream through the shared OfficeIMO client; byte-array output remains available when `Get-OfficeConfluenceAttachment` is called without `-OutFile`.

## Quick Start 🚀

Install and import the module once in the session:

```powershell
Install-Module -Name PSWriteOffice -Scope CurrentUser
Import-Module PSWriteOffice
```

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

### RTF

```powershell
New-OfficeRtf -OutputPath .\Report.rtf -Text 'Summary', 'Ready for review'
Get-OfficeRtf -Path .\Report.rtf
Update-OfficeRtfText -Path .\Report.rtf -OutputPath .\Report-Updated.rtf -OldText 'review' -NewText 'release'
ConvertFrom-OfficeRtf -Path .\Report-Updated.rtf -As Pdf -OutputPath .\Report.pdf
Get-OfficeDocumentChunk -Path .\Report-Updated.rtf
```

### Markdown and CSV

```powershell
New-OfficeMarkdown -Path .\README.md {
    MarkdownHeading -Level 1 -Text 'Report'
    MarkdownParagraph -Text 'Generated by PSWriteOffice.'
}

$data | ConvertTo-OfficeCsv -OutputPath .\export.csv
$rows | Export-OfficeCsv -Path .\pipes.csv -DelimiterText '||'
Import-OfficeCsv -Path .\pipes.csv -DelimiterText '||'
```

### Reader and Visio

```powershell
Get-OfficeDocumentCapability
Get-OfficeDocumentChunk -Path .\Report.docx
Get-OfficeDocumentChunk -Path .\Report.rtf
Get-OfficeDocument -Path .\Report.pdf -AsJson -Indented
$document = Get-OfficeDocument -Path .\Policy.docx -IncludePageLocations
$document | Search-OfficeDocument -Query 'retention period'
$document | Get-OfficeDocumentPageMarkdown -AsString
Get-OfficeDocumentAsset -Path .\Deck.pptx -Kind image -OutputDirectory .\reader-assets
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

```powershell
New-OfficeVisioGallery -OutputDirectory .\VisioGallery |
    Select-Object Name, FilePath, IsClean

Export-OfficeVisioStencilPreviewGallery -Path .\MyShapes.vssx -OutputDirectory .\StencilGallery -Title 'Custom stencil previews'
```

## Read, Inspect, and Convert 🔍

`PSWriteOffice` is not only about writing files. The module now has stronger read-back and bridge workflows too.

### Excel import/export fast path

```powershell
$data | Export-OfficeExcel -Path .\Report.xlsx -WorksheetName 'Data' -TableName 'Data' -AutoFit -FreezeTopRow
Import-OfficeExcel -Path .\Report.xlsx -WorksheetName 'Data'
```

### CSV, Excel, and database round trips

The examples below export SQL Server rows to CSV or Excel, import the file back as a `DataTable`, write it to a destination table with DbaClientX, and verify row counts.

```powershell
.\Examples\Csv\Example-CsvDbaClientXRoundTrip.ps1 -Server localhost -Database tempdb -RowCount 100
.\Examples\Excel\Example-ExcelDbaClientXRoundTrip.ps1 -Server localhost -Database tempdb -RowCount 100
```

### CSV and Excel benchmark snapshot

| Workflow | Rows | PSWriteOffice | Comparison | Result |
| --- | ---: | ---: | ---: | --- |
| Excel object table export | 10,000 | 301.2 ms | ImportExcel 2.03 s | 6.73x faster |
| Excel full-sheet import | 10,000 | 286.4 ms | ExcelFast 391.2 ms / ImportExcel 453.8 ms | Fastest |
| CSV to DataTable, mixed | 100,000 | 219.6 ms | NativeCsv 2.57 s | 11.72x faster |
| GZip CSV to DataTable, mixed | 10,000 | 77.5 ms | NativeCsv 205.2 ms | 2.65x faster |
| CSV object export, mixed | 100,000 | 209.0 ms | NativeCsv 225.2 ms | 8% faster |
| CSV object import, mixed | 100,000 | 471.7 ms | NativeCsv 554.4 ms | 18% faster |

The complete timing tables and commands are in [Benchmarks/README.md](Benchmarks/README.md).

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

### RTF bridges

```powershell
ConvertTo-OfficeRtf -WordPath .\Report.docx -OutputPath .\Report.rtf
ConvertTo-OfficeRtf -HtmlPath .\Report.html -OutputPath .\Report-from-html.rtf
ConvertFrom-OfficeRtf -Path .\Report.rtf -As Word -OutputPath .\Report-from-rtf.docx
ConvertFrom-OfficeRtf -Path .\Report.rtf -As Html -OutputPath .\Report.html -IncludeDefaultCss
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
- RTF now has create/read/update/replace cmdlets, Word/HTML/PDF bridges, and semantic Reader chunks through the `OfficeIMO.Reader.Rtf` adapter.
- `OfficeIMO.Reader` is surfaced through chunk, document-envelope, table, visual, asset, ingest, JSON, and capability-discovery cmdlets across the currently registered adapters.
- `OfficeIMO.Visio` is surfaced through a first diagram DSL plus built-in and external stencil catalogs, create/load/save, deterministic inspection, reference gallery generation, stencil preview gallery export, and SVG/PNG export cmdlets.
- Excel now has `Add-OfficeExcelTableOfContents`, `Get-OfficeExcelRange`, and `Get-OfficeExcelUsedRange`.
- PowerPoint now has section cmdlets, deck-wide text replacement, slide import helpers, slide copy, transitions, slide sizing, theme inspection, theme updates, and layout switching.
- Excel charts can now be finished with `Set-OfficeExcelChartLegend`, `Set-OfficeExcelChartDataLabels`, and `Set-OfficeExcelChartStyle`.
- Excel now has discoverable URL-image insertion plus smarter external hyperlink helpers.
- Excel summary sheets can now auto-link ranges and header-based columns to workbook tabs or external URLs.
- Backlink placement in Excel TOC flows is safer by default, avoiding overwriting active worksheet data.

## Documentation 📚

Use the examples in this README for the first pass, then move to the generated command reference when you need exact parameters and pipeline behavior.

- [Docs/Readme.md](Docs/Readme.md) is the generated cmdlet index.
- [Docs](Docs) contains one generated page per command.
- [Examples](Examples) contains runnable scripts grouped by Word, Excel, PowerPoint, PDF, RTF, Markdown, CSV, Visio, showcase scenarios, and shared samples.
- [Website/content/examples](Website/content/examples) contains the curated website examples for the project hub.

The generated command pages are built from XML comments in the cmdlet source. Prefer contextual examples there too: small workflow snippets with real input data, generated artifacts, readback/validation where useful, and no fake PowerShell compatibility aliases.

## Build and Test 🧪

```powershell
dotnet build .\Sources\PSWriteOffice.sln -c Debug
pwsh -NoLogo -NoProfile -File .\PSWriteOffice.Tests.ps1
pwsh -NoLogo -NoProfile -File .\Build\Build-Module.ps1 -RunMode Build
```

Regenerate the module manifest through the manifest lane:

```powershell
pwsh -NoLogo -NoProfile -File .\Build\Build-Module.ps1 -RunMode Manifest
```

Development loading is handled through [PSWriteOffice.psm1](PSWriteOffice.psm1), which prefers the local debug build in `Sources\PSWriteOffice\bin\Debug\`.

## License

MIT
