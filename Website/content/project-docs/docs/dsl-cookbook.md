---
title: "PSWriteOffice DSL Cookbook"
description: "Copy complete PowerShell recipes for Word, Excel, PowerPoint, PDF, Markdown, and multi-format document jobs."
layout: docs
---

The PSWriteOffice DSL is a set of scoped PowerShell commands for composing documents. The outer `New-Office*` command creates and saves the file. Nested blocks establish the active section, sheet, slide, page, or Markdown document so content commands do not need the document object repeated on every line.

## The common shape

```powershell
$rows = @(
    [pscustomobject]@{ Name = 'Alpha'; Status = 'Ready' }
    [pscustomobject]@{ Name = 'Beta'; Status = 'Review' }
)

New-OfficeWord -Path '.\Output\Status.docx' {
    WordSection {
        WordParagraph -Text 'Status' -Style Heading1
        WordTable -InputObject $rows
    }
}
```

The same plain PowerShell objects can feed an Excel table, PowerPoint chart, PDF table, or Markdown report. Keep source collection and business calculations outside the DSL. Let the composition block describe the artifact.

Canonical commands such as `Add-OfficeWordParagraph` are easiest to search in generated help and are a good default in shared scripts. Aliases such as `WordParagraph`, `ExcelTable`, `PptChart`, `PdfPanel`, and `MarkdownCallout` keep dense composition blocks readable. Both call the same cmdlets.

## Word recipes

- [Project status report](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Word/Recipe-Word-ProjectStatus.ps1) combines a header and footer, narrative, lists, conditional table rows, a chart, and approval controls.
- [Change approval checklist](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Word/Recipe-Word-ApprovalChecklist.ps1) creates a reusable form with a table of contents, content controls, and a watermark.
- [Executive report](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Showcase/Showcase-Word-ExecutiveReport.ps1) is the larger reference for metadata, bookmarks, footnotes, endnotes, tables, charts, and rich text.

Use Word when the reader needs a flowing, editable report with sections, review features, fields, or forms.

## Excel recipes

- [Project tracker](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Recipe-Excel-ProjectTracker.ps1) adds a structured table, status validation, conditional formatting, a chart, print settings, and an index sheet.
- [Budget dashboard](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Recipe-Excel-BudgetDashboard.ps1) separates summary formulas and charts from a styled detail table.
- [Operational dashboard](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Showcase/Showcase-Excel-OperationalDashboard.ps1) demonstrates KPI cells, tables, charts, pivots, sparklines, links, threaded comments, print layout, and workbook validation.

Use Excel when the data grid, formula model, filtering, chart interaction, or workbook navigation is part of the deliverable.

## PowerPoint recipes

- [Quarterly business review](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Recipe-PowerPoint-QuarterlyReview.ps1) creates a title slide, a data chart, a priority table, bullets, and speaker notes.
- [Training workshop](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Recipe-PowerPoint-TrainingWorkshop.ps1) builds learning objectives, an agenda table, a call-to-action slide, and presenter notes.
- [Service brief](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Showcase/Showcase-PowerPoint-ServiceBrief.ps1) combines semantic designer plans with direct slide composition, charts, sections, transitions, and inspection.

Use direct slide composition when the script owns placement. Use a deck plan when the content is semantic and the designer should choose layout variants.

## PDF recipes

- [Service invoice](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Recipe-Pdf-ServiceInvoice.ps1) composes invoice metadata, line items, totals, payment terms, headers, footers, and a link.
- [Audit report](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Recipe-Pdf-AuditReport.ps1) adds findings, bookmarks, page breaks, remediation sections, and interactive form fields.
- [Composed PDF report](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Example-PdfReportDsl.ps1) demonstrates themes, backgrounds, borders, rich text, rows, links, bookmarks, tables, and attachments.

Use PDF for fixed-layout delivery. Keep an editable source artifact as well when the workflow needs later content changes.

## Markdown recipes

- [Operations runbook](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Markdown/Recipe-Markdown-OperationsRunbook.ps1) creates front matter, a table of contents, warnings, task lists, code, a validation table, and collapsible rollback steps.
- [Release notes](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Markdown/Recipe-Markdown-ReleaseNotes.ps1) creates a release page with metadata, an upgrade callout, a change table, a checklist, code, and known limits.
- [Advanced Markdown](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Markdown/Example-MarkdownAdvanced.ps1) collects the broader typed Markdown surface in one script.

Use Markdown when the source should remain diffable, reviewable, and easy to publish into other text or document workflows.

## Create several formats from one data model

The [multi-format status pack](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Workflows/Recipe-MultiFormat-StatusPack.ps1) sends one service-status object array into five thin composition blocks:

```text
PowerShell objects
  |-- Markdown status page
  |-- Word owner report
  |-- Excel analysis workbook
  |-- PowerPoint review deck
  `-- PDF delivery copy
```

This is useful when audiences need different artifacts but the numbers and status labels must remain consistent. Calculate the data once, then keep each document block focused on how that audience consumes it.

## Run and adapt a recipe

```powershell
git clone https://github.com/EvotecIT/PSWriteOffice.git
Set-Location .\PSWriteOffice
pwsh .\Examples\Word\Recipe-Word-ProjectStatus.ps1 -OutputDirectory .\Output
```

Replace the sample objects first. Then adjust visual choices such as styles, colors, layout, and labels. Search the [PowerShell command reference](/api/powershell/) for exact parameters and accepted values.

The [complete example index](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/README.md) also covers Visio, Reader, RTF, CSV, HTML review, DbaClientX, ChartForgeX visuals, and Confluence publishing.
