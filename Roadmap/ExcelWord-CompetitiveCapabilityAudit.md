# Excel and Word Competitive Capability Audit

Date: 2026-05-16

This is the current competitive snapshot after the ImportExcel/PSWriteWord parity work.
Older backlog files were removed because the high-value wrapper work they tracked has
now shipped or moved into this smaller remaining-gap list.

## Summary

PSWriteOffice now covers the common ImportExcel muscle-memory paths without cloning
ImportExcel wholesale:

- `Export-OfficeExcel` handles pipeline/object export, append, clear-sheet, title,
  table styling, autofit, freeze, show, and `DataTable`/`DataSet`/`DataView`/`IDataReader`.
- `Import-OfficeExcel` covers worksheet/range/bounded reads and emits objects,
  hashtables, or `DataTable`.
- Worksheet copy, move, join, compare, print area/titles, find/replace, editable rows,
  chart axis/series/trendline formatting, comments, validation, conditional formatting,
  images, links, pivots, sparklines, and workbook summary inspection are wrapped.
- HTML and SQL-style ingestion are handled through bridge objects, not direct clients:
  `PSParseHTML`/HtmlTinkerX produce `DataTable`/`DataSet` or objects, and
  database modules such as DbaClientX/dbatools can feed those same shapes.

For Word, PSWriteOffice is already a much better replacement direction than
PSWriteWord. The old library remains useful mostly as inspiration for ergonomic
helpers: granular style commands, table mutation, image mutation, document merge,
page setup, equations, tab stops, and report-composer examples.

## Excel Status

| Capability | Current status | Remaining work |
| --- | --- | --- |
| One-shot object export | Wrapped by `Export-OfficeExcel` | Add only user-driven convenience switches after real migration feedback. |
| Import rows as objects | Wrapped by `Import-OfficeExcel` and lower-level range/table readers | Add coercion presets only if current raw/object/DataTable output proves insufficient. |
| Open/edit/save package flow | Wrapped by `Get-OfficeExcel`, `Save-OfficeExcel`, `Close-OfficeExcel`, and `Export-OfficeExcel -Append/-ClearSheet` | No ImportExcel-style package object clone needed. |
| Worksheet management | Copy, move, join, and range compare are wrapped | Remove/rename wrappers can be added if operators ask for them directly. |
| Tables and named ranges | Wrapped | Table totals and column-level table style knobs remain optional polish. |
| Formatting and styles | Rows, columns, formulas, headers/footers, conditional formatting, validation, print setup, gridlines, and chart finishing are wrapped | `Set-OfficeExcelColumnStyleByHeader` is the main remaining ergonomic gap. |
| Charts | Add chart plus axis, legend, labels, series, style, and trendline wrappers exist | Keep extending only where OfficeIMO exposes stable chart APIs. |
| Pivots and sparklines | Wrapped, but still need desktop Excel open-compatibility confidence before flagship examples rely on them | OfficeIMO engine compatibility tests are the right place for deeper fixes. |
| Find/replace and editable rows | Wrapped by `Find-OfficeExcel`, `Update-OfficeExcelText`, and `Edit-OfficeExcelRow` | Add more row-edit helpers only if maintenance scripts need them. |
| Print setup | Wrapped, including print area and titles | Covered. |
| Sheet/workbook protection | Sheet protect/unprotect is wrapped | Workbook encryption/password support should wait for real OfficeIMO engine support. |
| File conversion | CSV conversion exists | Range-to-image remains intentionally unowned until OfficeIMO has a renderer or we approve a rendering dependency. |
| HTML/data-source bridges | `Export-OfficeExcel` accepts bridge-friendly .NET data shapes; `Example-ExcelHtmlTablesViaPSParseHTML.ps1` shows HTML to Excel | Keep HTML parsing and SQL/OleDb clients outside PSWriteOffice core. |
| Diagnostics/schema | `Get-OfficeExcelSummary` exists | Schema inference is optional migration polish, not a current gap. |

## Word Status

| Capability | Current status | Remaining work |
| --- | --- | --- |
| Document create/load/save | Wrapped | Includes append/merge helpers for report packs. |
| Declarative report DSL | Word DSL aliases cover practical authoring | Do not port `Documentimo` verbatim; prefer one modern report-composer example. |
| Paragraph/text formatting | Core add/update/find/read helpers exist | Add compact `Set-OfficeWordRunStyle` and `Set-OfficeWordParagraphStyle` instead of many PSWriteWord-style micro-cmdlets. |
| Page setup | Wrapped | `Set-OfficeWordPageSetup` covers margins, size, orientation, and columns. |
| Tables | Object tables, table cells, conditional rows, nested tables, images/lists/chart anchoring are covered | Add row/column mutation, merge cell, layout, and width helpers. |
| Pictures/images | Basic image insertion exists | Add crop/fill, rotate, flip, transparency, fixed positioning, and alt-text mutation if OfficeIMO exposes them cleanly. |
| Headers/footers/page numbers | Wrapped | Covered. |
| Bookmarks/text replacement | Wrapped | Covered, with optional bookmark-text convenience later. |
| Document properties | Wrapped | Covered. |
| Charts | Wrapped | Add richer chart formatting only after command shapes stay consistent with Excel. |
| Equations and tab stops | Wrapped | `Add-OfficeWordEquation` and `Add-OfficeWordTabStop` expose the stable OfficeIMO.Word APIs. |
| TOC and fields | Wrapped | Covered. |
| Protection | Wrapped | Covered enough. |
| Cover pages | Wrapped | `Add-OfficeWordCoverPage` exposes template-driven cover pages with basic cover metadata. |
| Text boxes/shapes/SmartArt | Mostly missing | Start with predictable read/template helpers, not broad freeform authoring. |
| Comments/revisions/compare/statistics/macros/variables/embedded docs | Partially wrapped | `Get-OfficeWordStatistics` is wrapped; comments/compare should come before macros, and macros stay explicit/deferred. |

## Remaining Roadmap

### PowerShell Ergonomics

1. Add Excel column-style-by-header helpers for currency, percent, dates, durations,
   status fills, and common report formats.
2. Add a higher-level Excel report-sheet/composer wrapper over OfficeIMO's fluent
   reporting blocks.
3. Add Word run/paragraph style and table mutation helpers.
4. Add Word image layout and cover-page helpers for polished report output.

### OfficeIMO Engine First

1. Pivot and sparkline desktop-open compatibility confidence.
2. Workbook encryption/password support, if enterprise workbook parity becomes a goal.
3. Range-to-image only with an intentional renderer decision.
4. Richer Word image, shape, SmartArt, compare, and macro APIs only when the engine
   behavior is stable enough to expose safely.

### Explicitly Out of Core

- SQL/OleDb/web clients stay in data modules. PSWriteOffice consumes objects,
  `DataTable`, `DataSet`, `DataView`, and `IDataReader`.
- HTML parsing stays in PSParseHTML/HtmlTinkerX. PSWriteOffice should document the
  bridge and keep accepting the resulting data shapes.
- Excel COM automation is not a PSWriteOffice goal.
- PDF export and macros require explicit package/dependency/safety decisions.
