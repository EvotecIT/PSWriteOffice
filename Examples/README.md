# PSWriteOffice example library

These examples are complete PowerShell scripts that create or transform real files. Start with a focused recipe, then move to the larger showcase scripts when you want to see several DSL features working together.

```powershell
Install-Module PSWriteOffice -Scope CurrentUser
Import-Module PSWriteOffice
```

Most new recipes write to `Artefacts/Examples/<format>` so running them does not add generated documents to the source folders. Pass `-OutputDirectory` when you want the files somewhere else.

## Create documents with the DSL

| Format | Practical recipes | Larger examples |
| --- | --- | --- |
| Word | [Project status report](Word/Recipe-Word-ProjectStatus.ps1), [approval checklist](Word/Recipe-Word-ApprovalChecklist.ps1) | [Executive report](Showcase/Showcase-Word-ExecutiveReport.ps1), [advanced Word DSL](Word/Example-WordAdvanced.ps1) |
| Excel | [Project tracker](Excel/Recipe-Excel-ProjectTracker.ps1), [budget dashboard](Excel/Recipe-Excel-BudgetDashboard.ps1) | [Operational dashboard](Showcase/Showcase-Excel-OperationalDashboard.ps1), [advanced workbook](Excel/Example-ExcelAdvanced.ps1) |
| PowerPoint | [Quarterly review](PowerPoint/Recipe-PowerPoint-QuarterlyReview.ps1), [training workshop](PowerPoint/Recipe-PowerPoint-TrainingWorkshop.ps1) | [Service brief](Showcase/Showcase-PowerPoint-ServiceBrief.ps1), [themes and layouts](PowerPoint/Example-PowerPointThemeAndLayout.ps1) |
| PDF | [Service invoice](Pdf/Recipe-Pdf-ServiceInvoice.ps1), [audit report](Pdf/Recipe-Pdf-AuditReport.ps1) | [Composed PDF report](Pdf/Example-PdfReportDsl.ps1), [PDF operations](Pdf/Example-PdfOperations.ps1) |
| Markdown | [Operations runbook](Markdown/Recipe-Markdown-OperationsRunbook.ps1), [release notes](Markdown/Recipe-Markdown-ReleaseNotes.ps1) | [Advanced Markdown](Markdown/Example-MarkdownAdvanced.ps1), [Markdown DSL](Markdown/Example-MarkdownDsl.ps1) |
| Several formats | [One status pack from shared data](Workflows/Recipe-MultiFormat-StatusPack.ps1) | [Shared rich-text runs](Showcase/Showcase-RichTextRuns.ps1) |

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
