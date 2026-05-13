# Excel and Word Competitive Capability Audit

Date: 2026-05-13

Worktree: `C:\Support\GitHub\PSWriteOffice-excel-word-capability-audit`

Branch: `codex/excel-word-capability-audit`

## Sources

- PSWriteOffice command surface from `PSWriteOffice.psd1` and `Sources/PSWriteOffice/Cmdlets`.
- OfficeIMO engine source from sibling checkout `C:\Support\GitHub\OfficeIMO`. That checkout had unrelated local edits during this audit, so engine notes are a working-tree snapshot, not a released-package claim.
- ImportExcel `7.8.10`, current PSGallery version on 2026-05-13, saved locally under `Ignore/CapabilityAudit/Modules/ImportExcel/7.8.10`. Public source: <https://github.com/dfinke/ImportExcel>. Gallery package: <https://www.powershellgallery.com/packages/ImportExcel/7.8.10>.
- PSWriteWord `1.1.14`, current PSGallery version on 2026-05-13, saved locally under `Ignore/CapabilityAudit/Modules/PSWriteWord/1.1.14`, with source also available at `C:\Support\GitHub\PSWriteWord`. Public source: <https://github.com/EvotecIT/PSWriteWord>. Gallery package: <https://www.powershellgallery.com/packages/PSWriteWord/1.1.14>.

## Executive Summary

PSWriteOffice is already the right replacement direction. For Word it is more complete, more modern, and less tied to the old PS 5.1/Xceed DocX shape than PSWriteWord. The useful PSWriteWord leftovers are mostly ergonomics: very granular text formatting commands, table row/column mutation helpers, page setup wrappers, picture mutation, equation/tab-stop helpers, document merge examples, and the old `Documentimo` declarative report pattern.

For Excel, ImportExcel is still the benchmark for PowerShell operator ergonomics. PSWriteOffice/OfficeIMO already cover many underlying document features, but ImportExcel has a mature "pipe objects and get a useful workbook" surface with append, sheet clearing, pivot/chart-on-export, worksheet copy/merge/compare, query helpers, HTML import, range image export, and broad chart axis/series knobs. Some are simple PSWriteOffice wrappers over OfficeIMO. Some need OfficeIMO engine work first. Some should stay out of core scope because they are data-source bridges rather than document-format primitives.

The best path is not to clone every command name. Add a small set of high-value PSWriteOffice cmdlets that match PowerShell operator intent, and only push OfficeIMO where the engine must own real file-format behavior.

## Excel Crosswalk

| Capability | ImportExcel | PSWriteOffice today | OfficeIMO status | Recommendation |
| --- | --- | --- | --- | --- |
| One-shot object export | `Export-Excel` combines path, worksheet, table, style, title, freeze, autofit, chart, pivot, append, show, and password switches | `New-OfficeExcel` plus DSL primitives; powerful but more verbose for common reports | Engine supports sheets, tables, styles, autofit, charts, pivots, and load/save | Add `Export-OfficeExcel` as an operator cmdlet that composes the existing DSL. This is the biggest usability gap. |
| Import rows as objects | `Import-Excel` supports sheet/range bounds, headers, raw values, date/text coercion, passwords | `Get-OfficeExcelData`, `Get-OfficeExcelRange`, `Get-OfficeExcelUsedRange`, table/range readers | Typed reads, `RowsObjects()`, editable rows, header mapping, and read presets exist in OfficeIMO | Add `Import-OfficeExcel` as an alias-friendly read cmdlet with bounded range/header/coercion options. Keep `Get-*` readers for inspection. |
| Open/edit/save package flow | `Open-ExcelPackage`, `Close-ExcelPackage`, `Export-Excel -Append/-ClearSheet/-PassThru` | `Get-OfficeExcel`, `Save-OfficeExcel`, `Close-OfficeExcel`; no direct append/clear operator command | Load/edit/save exists | Add `Open-OfficeExcelPackage` only if needed, but prefer `Get-OfficeExcel` plus `Export-OfficeExcel -Append -ClearSheet -WorksheetName`. |
| Worksheet management | Add, copy, remove, select, merge, join, compare worksheets | Add sheet and sheet visibility are exposed; copy/remove/merge/compare are missing | Remove worksheet exists; rename exists internally; copy/merge/compare need confirmation or engine work | Add wrappers for remove/rename/reorder if stable. Add copy/merge/compare only after OfficeIMO has tested engine APIs. |
| Tables and named ranges | Tables, table styles, total settings, named ranges | Tables and named ranges are wrapped | Supported | Add table totals and more column-level table style knobs if OfficeIMO exposes them cleanly. |
| Formatting and styles | `Set-ExcelRange`, `Set-CellStyle`, `New-ExcelStyle`, `Set-ExcelColumn`, `Set-ExcelRow`, number formats | Cell/row/column, formulas, headers/footers, conditional formatting exist; general style builder is limited | OfficeIMO fluent `StyleBuilder`, column builders, header style helpers exist | Add `Set-OfficeExcelRangeStyle`, `Set-OfficeExcelColumnStyleByHeader`, and number-format presets. These should be wrappers first. |
| Charts | Chart definitions, pivot charts, axis titles/scale/number format, trendlines, legend, title, size | Add chart, legend, labels, style preset | OfficeIMO has rich `ExcelChart` APIs for axes, gridlines, display units, series colors, markers, trendlines, chart/plot area style | Add chart formatting cmdlets over existing engine APIs: axes first, then series/markers/trendlines. |
| Pivots | Pivot definitions, pivot charts, grouping, filters, totals, style | Pivot table creation exists; no pivot chart/grouping surface | Current OfficeIMO matrix calls pivots partial | Keep PSWriteOffice pivot creation, but do not overpromise. Engine needs parity tests before pivot chart/grouping wrappers. |
| Conditional formatting and validation | Conditional formatting builders, icon sets, data validation | Strong coverage already: rule, color scale, data bar, icon set, list/number/date/time/text/custom validation | Supported enough for wrappers | Low priority. Add convenience presets only after export ergonomics. |
| Comments | `Set-CellComment` | Add/remove comment exists | Supported | Covered. Maybe add `Set-OfficeExcelComment` alias shape if users expect mutation. |
| Find/replace and editable rows | ImportExcel can mutate values through package/range helpers | No explicit find/replace or editable row cmdlets | OfficeIMO has `FindFirst`, `ReplaceAll`, and `RowsObjects()` edit handles | Add `Find-OfficeExcelCell`, `Update-OfficeExcelText`, and `Edit-OfficeExcelRow`/`Set-OfficeExcelRowValue` wrappers. |
| Print setup | Freeze, print-ish layout through export knobs | Freeze, margins, orientation, page setup, gridlines | Print area/titles exist in OfficeIMO named-range APIs | Add `Set-OfficeExcelPrintArea` and `Set-OfficeExcelPrintTitles`. |
| Sheet/workbook protection | `Set-WorksheetProtection`; password parameters on open/export | Protect/unprotect sheet exists | Protection partial; encryption/password support is a roadmap gap | Expose more sheet protection options if stable. Password/encryption needs OfficeIMO engine work first. |
| File conversion | `ConvertTo-ExcelXlsx`, `Convert-ExcelRangeToImage` | CSV conversion exists, no range-to-image | Range-to-image is not visible in OfficeIMO.Excel | Keep CSV. Add range-to-image only if OfficeIMO grows a renderer or if we accept a rendering dependency. |
| HTML/data-source bridges | `Import-Html`, `Get-HtmlTable`, `Send-SQLDataToExcel`, `Read-OleDbData`, SQL insert conversion | Not present | Not core OfficeIMO responsibility | Defer. These belong in separate bridge modules or examples, not PSWriteOffice core. |
| Diagnostics/schema | `Get-ExcelFileSchema`, `Get-ExcelFileSummary`, workbook/sheet info | `Get-OfficeExcelSummary` exists | Inspection snapshots exist | Covered enough. Add schema inference only if real migration users need it. |

## Word Crosswalk

| Capability | PSWriteWord | PSWriteOffice today | OfficeIMO status | Recommendation |
| --- | --- | --- | --- | --- |
| Document create/load/save | `New/Get/Save/Merge-WordDocument` | `New/Get/Save/Close-OfficeWord`; no merge wrapper | OfficeIMO has document merge APIs | Add `Merge-OfficeWordDocument` and maybe `Append-OfficeWordDocument`; this is a real report-pack need. |
| Declarative report DSL | `Documentimo`, `DocText`, `DocTable`, `DocChart`, `DocToc`, list/page-break helpers | Word DSL aliases cover sections, paragraphs, tables, charts, TOC, breaks | Engine supports the underlying primitives | Do not port `Documentimo` verbatim. Add one `New-OfficeWordReport`/template-builder example if a more opinionated report DSL is desired. |
| Paragraph/text formatting | Many granular `Set-WordText*` commands for font, size, bold, italic, caps, hidden, highlight, kerning, language, spacing, underline, etc. | `Add-OfficeWordParagraph`, `Add-OfficeWordText`, update/find; styling is less granular from PowerShell | OfficeIMO paragraph/run style APIs exist | Add a compact `Set-OfficeWordRunStyle` and `Set-OfficeWordParagraphStyle` rather than dozens of micro-cmdlets. |
| Page setup | Margins, orientation, page size, page settings | Sections exist, but page setup wrappers are thin/missing | OfficeIMO section/page setup APIs exist | Add `Set-OfficeWordPageSetup` covering margins, size, orientation, columns if stable. |
| Tables | Row/column add/remove/copy, borders, cell color/shading, direction, auto-fit, widths, merge cells | Object tables, table cell DSL, conditional formatting; row/column mutation and merge wrappers are limited | OfficeIMO table APIs include merge cells, layout, widths, styles, comments | Add `Merge-OfficeWordTableCell`, `Set-OfficeWordTableLayout`, `Set-OfficeWordTableColumnWidth`, row/column add/remove wrappers. |
| Pictures/images | Add/get/set/remove picture with rotation, flip, description, dimensions | Add image with width/height/wrap/description only | OfficeIMO image/wrap/location APIs exist; richer image style exists | Add image mutation cmdlets for rotate, flip, crop/fill mode, transparency, fixed positioning, and alt text if OfficeIMO exposes them cleanly. |
| Headers/footers/page numbers | Add/get header/footer, page number/count | Wrapped | Supported | Mostly covered. Add page-count style if needed. |
| Bookmarks/text replacement | Bookmark get/set and text replace examples | Bookmark, field, hyperlink, find/update text are wrapped | Supported | Covered. Consider `Set-OfficeWordBookmarkText` convenience if migration users ask. |
| Document properties | Custom property get/add | Built-in/custom document properties wrapped | Supported | Covered. |
| Charts | Bar/line/pie plus series | Word chart wrapper exists | Supported enough for reports | Add series/axis formatting only after Excel chart work, so command shape stays consistent. |
| Equations and tab stops | `Add-WordEquation`, `Add-WordTabStopPosition` | Missing | OfficeIMO has equation and tab-stop primitives | Add wrappers. These are small and useful for parity. |
| TOC | Add TOC and TOC item helpers | Add/set/get/update/remove TOC | Supported | Covered. |
| Protection | Add protection | `Protect-OfficeWordDocument` | Supported | Covered enough. |
| Cover pages | Not a major PSWriteWord surface | Missing | OfficeIMO has cover-page templates | Add `Add-OfficeWordCoverPage` from OfficeIMO templates. High visual value. |
| Text boxes/shapes/SmartArt | Not strong in PSWriteWord; OfficeIMO has newer features | Missing | OfficeIMO has text boxes, shapes, SmartArt | Add wrappers after image/page/table polish. SmartArt should begin as template/read-safe helpers, not broad freeform authoring. |
| Comments/revisions/compare/statistics/macros/variables/embedded docs | Mostly not the old PSWriteWord core | Mostly missing | OfficeIMO has comment/revision/statistics/macro/variable/embedded/comparer related APIs | Treat as a second wave. Add comments/statistics/compare before macros. Keep macros explicit/deferred. |

## Recommended Backlog

### Phase 1: PowerShell Ergonomics, Mostly Wrappers

1. Add `Export-OfficeExcel` for the common ImportExcel-style pipeline scenario.
2. Add `Import-OfficeExcel` with sheet/range/header/coercion options and map it to the existing read services.
3. Add `Set-OfficeExcelRangeStyle` and `Set-OfficeExcelColumnStyleByHeader` with number-format presets.
4. Add `Find-OfficeExcelCell` and `Update-OfficeExcelText` over OfficeIMO `FindFirst`/`ReplaceAll`.
5. Add `Set-OfficeExcelPrintArea` and `Set-OfficeExcelPrintTitles`.
6. Add Word wrappers for `Merge-OfficeWordDocument`, `Set-OfficeWordPageSetup`, table cell merge/layout/widths, equation, and tab stops.

### Phase 2: OfficeIMO Engine Parity First

1. Excel worksheet copy, reorder, join, merge, and compare with compatibility tests.
2. Excel pivot chart/grouping/filter parity tests before exposing more pivot parameters.
3. Excel password/encryption if we want enterprise workbook parity with ImportExcel/EPPlus expectations.
4. Range-to-image only if OfficeIMO owns or intentionally depends on a renderer.
5. Word image layout mutation beyond simple inline/wrap: crop/fill, rotate, flip, absolute positioning, and reliable alt text.

### Phase 3: Higher-Level Polish

1. A PSWriteOffice report-composer layer that uses OfficeIMO `SheetComposer` and Word fluent builders rather than copying `Documentimo`.
2. Word cover page wrappers and reusable executive-report templates.
3. Word comments, document statistics, compare, variables, and embedded-document helpers.
4. Consistent chart formatting commands across Excel, Word, and PowerPoint where the engines support comparable concepts.

## What Not To Pull Into Core Yet

- ImportExcel's SQL/OleDb/UPS/USPS/HTML scraping helpers should remain external bridge examples unless there is a clear PSWriteOffice ownership decision.
- `Documentimo` should be treated as inspiration for an opinionated report DSL, not as a compatibility target.
- Macros and PDF export should stay explicit/deferred because they change dependency and safety expectations.
- Password/encryption should not be exposed as a hollow parameter until OfficeIMO can actually honor it.

## Immediate Next PR Candidate

The most useful first implementation PR is `Export-OfficeExcel` plus a small ImportExcel migration example:

- `Export-OfficeExcel -Path .\Report.xlsx -WorksheetName Data -InputObject $rows -TableName Data -AutoFit -FreezeTopRow -BoldTopRow -Show`
- optional `-Append`, `-ClearSheet`, `-Title`, `-TableStyle`, `-NoNumberConversion`, `-NoHyperlinkConversion`
- optional simple chart and pivot switches can come after the base export path is stable

This would give users the common ImportExcel muscle-memory path while still keeping OfficeIMO as the document engine and PSWriteOffice as the PowerShell-friendly layer.
