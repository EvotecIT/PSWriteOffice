# PSWriteOffice Showcase and OfficeIMO Polish Plan

Date: 2026-05-16

This plan tracks the current showcase and polish work after the broad OfficeIMO
wrapper pass. Completed backlog slices were removed; this file now focuses only on
what still moves the product forward.

## Current Position

PSWriteOffice is no longer missing the common primitives.

- Word wraps document lifecycle, readers, paragraphs, lists, tables, table-cell
  content, conditional table formatting, TOC, bookmarks, fields, footnotes/endnotes,
  content controls, charts, hyperlinks, document properties, backgrounds, watermarks,
  protection, mail merge, HTML conversion, and Markdown conversion.
- Excel wraps import/export, sheets, cells, rows, columns, tables, `DataTable` and
  `DataSet` ingestion, named ranges, formulas, validation, conditional formatting,
  comments, images and URL images, charts and chart finishing, pivots, sparklines,
  TOC/navigation, internal links, URL links, smart hyperlinks, print setup, header
  and footer images, gridlines, freeze panes, sheet visibility, sorting, autofit,
  worksheet copy/move/join/compare, find/replace, editable rows, range/read helpers,
  and workbook summary inspection.
- PowerPoint wraps deck lifecycle, slides, titles, text boxes, bullets, notes,
  sections, tables, images, shapes, charts, backgrounds, layouts, layout placeholders,
  layout boxes, theme colors/fonts/name, slide transitions, slide sizing, slide import,
  slide copy, text replacement, inspection helpers, and the initial OfficeIMO
  designer/deck-plan bridge.

The next step is opinionated composition: fewer coordinates, better defaults, richer
diagnostics, and showcase examples that feel like real reporting products.

## Remaining Product Gaps

### Word

Focus on professional report assembly:

1. Image layout wrappers for crop/fill, rotation, transparency, fixed positioning,
   wrapping, and alt text.
2. Cover-page and append/merge-document helpers.
3. Page setup and compact run/paragraph style helpers.
4. Table row/column mutation, merge-cell, layout, and width helpers.
5. Equation and tab-stop wrappers.

Keep macros, SmartArt authoring, PDF export, compare, and embedded-document work as
explicit scope decisions.

### Excel

Focus on human-readable workbooks and migration ergonomics:

1. `Add-OfficeExcelReportSheet` / `ExcelReportSheet` wrapper over OfficeIMO's fluent
   report blocks.
2. `Set-OfficeExcelColumnStyleByHeader` for currency, percentages, dates, durations,
   and status fills without range math.
3. KPI, legend, callout, section, and reference blocks for dashboard composition.
4. Execution policy and diagnostics only if they remain simple from PowerShell.
5. OfficeIMO engine confidence for pivot and sparkline desktop-open compatibility.

Keep SQL/OleDb clients, HTML parsing, Excel COM, workbook passwords, and range-to-image
outside core until there is an explicit ownership/dependency decision.

### PowerPoint

Focus on semantic decks and layout safety:

1. Metrics and visual-frame deck-plan helpers.
2. Richer recommendation and content-fit diagnostics.
3. Shape layout commands for align, distribute, stack, grid, fit-to-bounds, resize,
   z-order, duplicate, and group.
4. Richer chart formatting matching Excel concepts where practical.
5. Table cell formatting helpers for padding, borders, row heights, merged cells, and
   preset styles.
6. Slide hidden/reorder controls.

## Showcase Examples

The three flagship examples exist and should remain the main regression/demo path:

- `Examples/Showcase/Showcase-Word-ExecutiveReport.ps1`
- `Examples/Showcase/Showcase-Excel-OperationalDashboard.ps1`
- `Examples/Showcase/Showcase-PowerPoint-ServiceBrief.ps1`

### Word Showcase

Current scenario: executive service-health report generated from PowerShell objects.

Keep demonstrating:

- header/footer, page numbers, document properties
- TOC and heading hierarchy
- executive summary and action sections
- status table with conditional formatting
- chart section
- hyperlink/bookmark navigation
- content controls for approvals
- watermark/background
- footnote and endnote

Next polish:

- cover page helper once stable
- richer image layout/alt text
- merge/append-document appendix flow

### Excel Showcase

Current scenario: multi-sheet operational workbook with summary, details, trend, and
appendix tabs.

Keep demonstrating:

- summary dashboard sheet
- TOC with links and backlinks
- object tables with friendly headers
- conditional color scale/data bars/icon sets
- validation list for status
- owner summary sheet while pivot compatibility is proven
- charts with styled labels/legend/axes/series
- internal links by header
- smart external links
- URL/local image
- print setup and header/footer logo
- hidden notes/config sheet
- `Get-OfficeExcelSummary` validation

Next polish:

- report composer wrapper over OfficeIMO `SheetComposer`
- column-style-by-header wrapper
- pivot/sparkline compatibility confidence before relying on them in flagship output

### PowerPoint Showcase

Current scenario: service/consulting deck with executive story, process, KPIs,
coverage, proof points, charts, and appendix.

Keep demonstrating:

- branded theme and fonts
- slide size and transitions
- sections
- title/section slide
- process slide
- card grid slide
- coverage/proof slide
- chart slide
- table slide
- image/background slide
- speaker notes
- imported/copy slide appendix
- inspection summary after save

Next polish:

- metric strip and visual-frame semantic helpers
- shape layout commands
- chart formatting wrappers
- table cell formatting wrappers
- fit diagnostics that explain layout choices

## Blog Series

Create three separate posts in `C:\Support\GitHub\Website.Contributions` once
screenshots are produced from the generated artifacts.

| Post | Working title | Required assets |
| --- | --- | --- |
| Word | Build a polished Word executive report from PowerShell | generated DOCX, cover/hero image, screenshots of cover/TOC/chart pages, code excerpt |
| Excel | Build an Excel operational dashboard from PowerShell | generated XLSX, screenshots of summary/detail/chart pages, code excerpt |
| PowerPoint | Build a beautiful PowerPoint service brief from PowerShell | generated PPTX, screenshots of title/process/metric/chart slides, code excerpt |

Blog visuals should show real generated output. Use generated imagery only for covers
or explanatory graphics, not as a substitute for screenshots.

## Visual Pipeline

1. Generate DOCX/XLSX/PPTX examples into `Examples/Documents`.
2. Export representative pages/slides/sheets to PNG/WebP.
3. Store article images under each post's image folder.
4. Link generated Office artifacts from the article or release assets when possible.

Possible export paths:

- PowerPoint: installed Office PowerPoint if available, otherwise LibreOffice/headless
  fallback.
- Word: installed Word or future `OfficeIMO.Word.Pdf` if that package scope is approved.
- Excel: installed Excel for selected sheets, or a future HTML/PNG preview helper.

## Acceptance Criteria

- Each product has one flagship showcase script that creates a non-trivial artifact
  from PowerShell objects.
- Each artifact includes navigation, visual hierarchy, structured data, and at least
  one chart or visual.
- Each example is fast enough to run as a smoke test and deterministic enough to
  inspect in CI.
- Each blog post uses runnable code plus real screenshots of generated output.
- Missing capability is fixed at the lowest sensible layer: OfficeIMO for engine
  behavior, PSWriteOffice for PowerShell ergonomics.
