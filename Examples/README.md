# PSWriteOffice example library

These examples are complete PowerShell scripts that create or transform real files. Start with a focused recipe, then move to the larger showcase scripts when you want to see several DSL features working together.

```powershell
Install-Module PSWriteOffice -Scope CurrentUser
```

Each `Recipe-*` script uses simple paths such as `.\Project-Status.docx`. Run a recipe from the folder where you want its documents, or copy the composition block into your own script and change only the file names.

Creation blocks in the `Recipe-*` scripts use the short DSL aliases consistently, including `WordNew`, `ExcelNew`, `PptNew`, `PdfNew`, and `MarkdownNew`. Recipes that read, update, merge, split, or convert existing files use canonical cmdlet names so those operations are easy to discover in command help. The [DSL cookbook](https://officeimo.com/docs/pswriteoffice/dsl-cookbook/) shows composition in alias and canonical forms so you can choose one style for your own script.

Saved DSL constructors are quiet by default. Add `-PassThru` only when the next command needs the saved file.

## Create documents with the DSL

| Format | Practical recipes | Larger examples |
| --- | --- | --- |
| Word | [Project status report](Word/Recipe-Word-ProjectStatus.ps1), [approval checklist](Word/Recipe-Word-ApprovalChecklist.ps1) | [Executive report](Showcase/Showcase-Word-ExecutiveReport.ps1), [advanced Word DSL](Word/Example-WordAdvanced.ps1) |
| Excel | [Project tracker](Excel/Recipe-Excel-ProjectTracker.ps1), [budget dashboard](Excel/Recipe-Excel-BudgetDashboard.ps1) | [Operational dashboard](Showcase/Showcase-Excel-OperationalDashboard.ps1), [advanced workbook](Excel/Example-ExcelAdvanced.ps1) |
| PowerPoint | [Quarterly review](PowerPoint/Recipe-PowerPoint-QuarterlyReview.ps1), [training workshop](PowerPoint/Recipe-PowerPoint-TrainingWorkshop.ps1) | [Service brief](Showcase/Showcase-PowerPoint-ServiceBrief.ps1), [themes and layouts](PowerPoint/Example-PowerPointThemeAndLayout.ps1) |
| PDF | [Service invoice](Pdf/Recipe-Pdf-ServiceInvoice.ps1), [audit report](Pdf/Recipe-Pdf-AuditReport.ps1) | [Composed PDF report](Pdf/Example-PdfReportDsl.ps1), [PDF operations](Pdf/Example-PdfOperations.ps1) |
| Markdown | [Operations runbook](Markdown/Recipe-Markdown-OperationsRunbook.ps1), [release notes](Markdown/Recipe-Markdown-ReleaseNotes.ps1) | [Advanced Markdown](Markdown/Example-MarkdownAdvanced.ps1), [Markdown DSL](Markdown/Example-MarkdownDsl.ps1) |
| Several formats | [One status pack from shared data](Workflows/Recipe-MultiFormat-StatusPack.ps1) | [Shared rich-text runs](Showcase/Showcase-RichTextRuns.ps1) |

## Read, modify, combine, and convert

| Format | Read and inspect | Modify or combine | Convert or deliver |
| --- | --- | --- | --- |
| Word | [Inspect an existing document](Word/Recipe-Word-InspectExisting.ps1) | [Update existing content](Word/Recipe-Word-UpdateExisting.ps1), [merge documents](Word/Recipe-Word-MergeDocuments.ps1), [mail-merge letters](Word/Recipe-Word-MailMergeLetters.ps1) | [Word and Markdown conversion](Word/Example-WordMarkdownConvert.ps1), [HTML review](Word/Example-WordHtmlConvert.ps1) |
| Excel | [Read and filter rows](Excel/Recipe-Excel-ReadAndFilter.ps1) | [Update an existing workbook](Excel/Recipe-Excel-UpdateExisting.ps1), [merge workbooks](Excel/Recipe-Excel-MergeWorkbooks.ps1), [compare workbooks](Excel/Recipe-Excel-CompareWorkbooks.ps1) | [Import delimited data](Excel/Recipe-Excel-ImportDelimited.ps1), [HTML review](Excel/Example-ExcelHtmlReview.ps1) |
| PowerPoint | [Inspect a deck](PowerPoint/Recipe-PowerPoint-InspectDeck.ps1) | [Update existing content](PowerPoint/Recipe-PowerPoint-UpdateExisting.ps1), [reuse and combine slides](PowerPoint/Recipe-PowerPoint-ReuseSlides.ps1) | [HTML review](PowerPoint/Example-PowerPointHtmlReview.ps1) |
| PDF | [Inspect and preflight](Pdf/Recipe-Pdf-InspectAndPreflight.ps1) | [Merge and split](Pdf/Recipe-Pdf-MergeAndSplit.ps1), [position content](Pdf/Recipe-Pdf-PositionedCanvas.ps1), [redact detected text](Pdf/Recipe-Pdf-RedactDetectedText.ps1) | [Forms](Pdf/Recipe-Pdf-Forms.ps1), [sanitize and optimize](Pdf/Recipe-Pdf-SanitizeAndOptimize.ps1) |
| Markdown | [Inspect structured content](Markdown/Recipe-Markdown-InspectContent.ps1) | [Convert to and from Word](Markdown/Recipe-Markdown-WordRoundTrip.ps1) | [Publish HTML](Markdown/Recipe-Markdown-PublishHtml.ps1) |
| Reader | [Search a mixed folder](Reader/Recipe-Reader-SearchFolder.ps1) | [Extract chunks and tables](Reader/Recipe-Reader-ExtractTables.ps1) | [Ingest a bounded folder](Reader/Recipe-Reader-IngestFolder.ps1) |

## Inspect, convert, and integrate

- [Visio architecture map](Visio/Example-Visio-ArchitectureMap.ps1) and [network topology](Visio/Example-Visio-NetworkTopology.ps1)
- [Mixed-document search](Reader/Example-MixedDocumentSearch.ps1)
- [RTF and Markdown round trip](Rtf/Example-RtfMarkdownRoundTrip.ps1)
- [CSV basics](Csv/Example-CsvBasic.ps1), [advanced CSV options](Csv/Example-CsvAdvanced.ps1), and [DbaClientX round trip](Csv/Example-CsvDbaClientXRoundTrip.ps1)
- [Excel and DbaClientX round trip](Excel/Example-ExcelDbaClientXRoundTrip.ps1)
- [HTML review for Word](Word/Example-WordHtmlConvert.ps1), [Excel](Excel/Example-ExcelHtmlReview.ps1), and [PowerPoint](PowerPoint/Example-PowerPointHtmlReview.ps1)
- [ChartForgeX visuals](Visuals/Example-ChartForgeXVisuals.ps1)
- [Confluence report publishing](Confluence/Example-ConfluenceAzureTableReport.ps1)

The [OfficeIMO website](https://officeimo.com/docs/pswriteoffice/) ingests these scripts for the PowerShell reference and publishes the authored PSWriteOffice guides from this repository. Command pages remain the source of truth for exact parameters; the recipes show how those commands fit into complete jobs.
