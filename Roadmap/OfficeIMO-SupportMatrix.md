# PSWriteOffice and OfficeIMO Support Matrix

Date: 2026-06-11

This matrix is the current planning companion for `OfficeIMO-Showcase-PolishPlan.md`.
For the ImportExcel and PSWriteWord competitive crosswalk, see
`ExcelWord-CompetitiveCapabilityAudit.md`.

## Legend

- `Wrapped`: available through PSWriteOffice today.
- `Partial wrapper`: a useful surface exists, but deeper OfficeIMO capability remains
  to expose or prove.
- `Wrapper gap`: OfficeIMO appears to have a useful capability that PSWriteOffice
  should expose.
- `Engine gap`: needs OfficeIMO work first or upstream API stabilization.
- `Deferred`: useful, but outside core scope unless approved.

## Word

| Capability | Status | Notes |
| --- | --- | --- |
| Create/load/save documents | Wrapped | `New-OfficeWord`, `Get-OfficeWord`, `Save-OfficeWord`, `Close-OfficeWord`, including encrypted package open/save |
| Sections, headers, footers, page numbers | Wrapped | Good enough for report examples |
| Paragraphs, text, lists | Partial wrapper | Core authoring is wrapped; richer run/paragraph style builders remain |
| Tables from objects and table-cell content | Partial wrapper | Object tables, conditional rows, nested tables, lists, images, chart anchoring, table-cell read/style, width, merge, and split helpers are wrapped; row/column mutation remains |
| TOC and field updates | Wrapped | `Add/Set/Get/Update/Remove-OfficeWordTableOfContent`, `Update-OfficeWordFields` |
| Bookmarks and hyperlinks | Wrapped | External and anchor links are exposed |
| Document properties | Wrapped | Built-in/custom property surface exists |
| Backgrounds and watermarks | Wrapped | Color/image backgrounds and watermarking exist |
| Content controls | Wrapped | Checkbox, date, dropdown, combo, picture, repeating section |
| Charts | Wrapped | Good enough for showcase reports |
| HTML and Markdown conversion | Wrapped | Useful for sidecar previews/blog code |
| Mail merge | Wrapped | Suitable for practical examples |
| Footnotes/endnotes | Wrapped | Add/read wrappers return document-safe note snapshots |
| Page setup and columns | Wrapped | `Set-OfficeWordPageSetup` covers page size, orientation, margins, and columns |
| Advanced image layout | Partial wrapper | `Get/Set-OfficeWordImage` exposes crop, rotation, flip, wrapping, metadata, and visibility; fixed-position semantics remain engine-led |
| Text boxes and shapes | Partial wrapper | `Add/Get/Set-OfficeWordShape` exposes basic shape authoring and styling; text boxes and richer templates remain |
| Cover pages | Wrapped | `Add-OfficeWordCoverPage` exposes stable OfficeIMO templates and basic cover metadata |
| Append/merge documents | Wrapped | `Join-OfficeWordDocument` appends one or more documents into a base document |
| Equations and tab stops | Wrapped | `Add-OfficeWordEquation` and `Add-OfficeWordTabStop` expose stable OfficeIMO.Word APIs |
| Document statistics | Wrapped | `Get-OfficeWordStatistics` exposes page/paragraph/word/object counts |
| Macros | Deferred | Keep preview-only if added |
| SmartArt authoring | Deferred | Detection/read helpers are safer first |
| PDF export | Wrapped | `New/Save-OfficeWord -PdfPath` use OfficeIMO.Word.Pdf sidecar export |

## Excel

| Capability | Status | Notes |
| --- | --- | --- |
| Create/load/save workbook | Wrapped | `New/Get/Save/Close-OfficeExcel`, including encrypted package open/save |
| Import/export operator flow | Wrapped | `Export-OfficeExcel` and `Import-OfficeExcel` cover common ImportExcel-style workflows |
| Bridge data shapes | Wrapped | Objects, dictionaries, `DataTable`, `DataSet`, `DataView`, and `IDataReader` are accepted by table/export paths; performance selection stays inside OfficeIMO normal APIs |
| Sheets, cells, rows, columns | Wrapped | Strong primitive coverage |
| Worksheet copy/move/join/compare | Wrapped | Useful maintenance and migration helpers exist |
| Tables from objects | Wrapped | Core reporting path works |
| Named ranges | Wrapped | Includes hidden and validation-mode support |
| TOC and workbook navigation | Wrapped | Includes internal links/backlinks |
| Range/data/table/pivot readers | Wrapped | Good read-back foundation |
| Validation | Wrapped | List, whole number, decimal, date, time, text length, custom formula |
| Conditional formatting | Wrapped | Rules, color scale, data bar, icon set |
| Charts | Wrapped | Add chart plus axis, legend, data labels, style, series, and trendline helpers |
| Pivot tables and sparklines | Partial wrapper | Cmdlets exist, but desktop Excel open compatibility needs engine confidence before flagship examples rely on them |
| Images and URL images | Wrapped | Includes in-sheet and header/footer image paths |
| Hyperlinks | Wrapped | Raw, smart, host, internal, URL by header |
| Print setup | Wrapped | Orientation, margins, page setup, gridlines, freeze, print area, print titles |
| Sort/autofit | Wrapped | Direct helpers exist |
| Find/replace and editable rows | Wrapped | `Find-OfficeExcel`, `Update-OfficeExcelText`, `Edit-OfficeExcelRow` |
| Workbook summary inspection | Wrapped | `Get-OfficeExcelSummary` reports workbook shape and major object collections |
| HTML to Excel bridge | Wrapped by example | `Example-ExcelHtmlTablesViaPSParseHTML.ps1` keeps HTML parsing outside PSWriteOffice |
| Fluent report composer | Wrapped | `Add-OfficeExcelReportSheet` exposes OfficeIMO `SheetComposer` as a PowerShell report sheet DSL |
| KPI, legend, callout, table blocks | Wrapped | Report-sheet cmdlets cover the first reusable dashboard blocks |
| Column style by header | Wrapped | `Set-OfficeExcelColumnStyleByHeader` handles currency, percent, dates, durations, fills, and status maps without range math |
| Execution policy/diagnostics | Wrapped | `Set-OfficeExcelExecutionPolicy` plus save preflight/repair/validation switches expose the simple OfficeIMO knobs |
| Workbook encryption/passwords | Wrapped | OfficeIMO encrypted package APIs are surfaced through lifecycle commands |
| Range to image | Deferred | Explicitly out of parity scope unless OfficeIMO intentionally grows a pure renderer |
| Google Sheets bridge | Deferred | Explicit package-scope expansion only |
| PDF export | Wrapped | `New/Save-OfficeExcel -PdfPath` use OfficeIMO.Excel.Pdf sidecar export |

## PowerPoint

| Capability | Status | Notes |
| --- | --- | --- |
| Create/load/save decks | Wrapped | `New/Get/Save/Close-OfficePowerPoint`, including encrypted package open/save |
| Slides, title, text box, bullets | Wrapped | Current examples use these |
| Tables, images, shapes | Wrapped | Primitive coverage exists |
| Charts | Wrapped | Column, pie, doughnut, scatter exposed |
| Backgrounds | Wrapped | Color and image background |
| Notes | Wrapped | Read/write speaker notes |
| Sections | Wrapped | Add/list/rename |
| Themes | Wrapped | Theme color/font/name |
| Layouts/placeholders/layout boxes | Wrapped | Stronger than baseline |
| Transitions and slide size | Wrapped | Good for showcase deck polish |
| Import/copy slides | Wrapped | Useful appendix workflow |
| Text replacement and inspection | Wrapped | Good maintenance surface |
| Designer brief/recipes/directions | Partial wrapper | `Add-OfficePowerPointDesignerDeck` exposes the first high-level bridge |
| Deck plan and semantic slides | Partial wrapper | Section, process, card grid, coverage, capability, case study, and logo wall helpers exist; metrics and visual frames remain |
| Content-fit diagnostics | Partial wrapper | Preview/render summaries exist; richer fit warnings should still be exposed |
| Shape layout helpers | Wrapper gap | Align, distribute, stack, grid, fit, resize, z-order, duplicate, group |
| Guides and snap-to-grid | Wrapper gap | Useful for manual layout scripts |
| Rich chart formatting | Wrapper gap | Match OfficeIMO chart title/legend/axis/series/marker/trendline APIs |
| Table cell formatting | Wrapper gap | Merged cells, padding, borders, row heights, autofit, preset style wrappers |
| Slide hidden/reorder controls | Wrapper gap | Useful deck maintenance commands |
| PDF export | Wrapped | `New/Save-OfficePowerPoint -PdfPath` use OfficeIMO.PowerPoint.Pdf sidecar export |

## PDF

| Capability | Status | Notes |
| --- | --- | --- |
| Create/load/save documents | Wrapped | `New/Get/Save-OfficePdf` wrap OfficeIMO.Pdf directly |
| Composition blocks | Wrapped | Themes, headings, paragraphs, rich inline text/link runs, lists, tables, images, panels, horizontal rules, bookmarks, spacers, row/column layout, page breaks, headers, footers, metadata, watermarks, backgrounds, background images/shapes, page borders, page setup, font options, attachments, and form fields |
| Existing-PDF readback | Wrapped | Info, preflight, text, logical Markdown, form fields, images, and attachments |
| Existing-PDF operations | Wrapped | Join, split, copy, remove, move, rotate through `Set-OfficePdfPage -Rotation`, metadata updates, form fill/flat conversion, and text/image stamps |
| Compliance readiness | Partial wrapper | Generated document profile/groundwork and readiness reports are wrapped; existing-PDF conformance checks remain engine-led |
| HTML/PDF conversion | Wrapped | `ConvertFrom-OfficePdfHtml` and `ConvertTo-OfficePdfHtml` expose OfficeIMO.Html.Pdf semantic/document and semantic/positioned-review profiles |
| Signatures, encryption, redaction | Engine gap | Add only when OfficeIMO.Pdf exposes stable reusable APIs |

## Reader

| Capability | Status | Notes |
| --- | --- | --- |
| Capability discovery | Wrapped | `Get-OfficeDocumentCapability` lists built-in and modular handlers after registering the OfficeIMO.Reader.Pdf adapter |
| Chunk extraction | Wrapped | `Get-OfficeDocumentChunk` wraps `DocumentReader.Read` and `ReadFolder` for Word, Excel, PowerPoint, Markdown, PDF, and text-like files |
| Document envelope | Wrapped | `Get-OfficeDocument` returns `OfficeDocumentReadResult` or deterministic JSON through `-AsJson` |
| PDF reader adapter | Wrapped | `OfficeIMO.Reader.Pdf` is referenced and registered so the modular PDF handler can replace the built-in PDF capability |
| Tables, assets, and visuals | Wrapper gap | OfficeIMO.Reader exposes deeper table/asset/visual APIs; PSWriteOffice currently surfaces the chunk and document-envelope workflows first |

## Visio

| Capability | Status | Notes |
| --- | --- | --- |
| Create/load/save documents | Wrapped | `New/Get/Save-OfficeVisio` expose the core `.vsdx` lifecycle |
| Basic diagram DSL | Wrapped | `New-OfficeVisio { VisioRectangle ...; VisioConnector ... }` covers pages, rectangles, ellipses, diamonds, text boxes, connectors, and stencil shapes |
| Inspection snapshots | Wrapped | `Get-OfficeVisioInfo` exposes deterministic `CreateInspectionSnapshot()` output and stable text |
| SVG/PNG export | Wrapped | `ConvertTo-OfficeVisioSvg` and `ConvertTo-OfficeVisioPng` expose dependency-free OfficeIMO.Visio renderers |
| Stencil catalogs | Wrapped | `Get/Find/Import-OfficeVisioStencil*` expose built-in catalogs, installed/package-backed catalog loading, search, and `VisioStencil` placement |
| Shape and connector authoring | Partial wrapper | Basic shape/connectors/stencil DSL exists; grouping, layers, layout helpers, and semantic diagram builders remain future wrapper slices |
| Semantic diagram builders | Wrapper gap | Keep this in OfficeIMO.Visio first, then expose focused PowerShell workflows once the high-level API is stable |

## Recommended Next PRs

1. Word run/paragraph style, row/column table mutation, and text box helpers.
2. PowerPoint metrics/visual-frame helpers, fit diagnostics, and shape layout polish.
3. Visio grouping/layer/layout and semantic diagram builder wrappers after OfficeIMO.Visio stabilizes the high-level workflows.
4. OfficeIMO engine confidence for Excel pivot/sparkline desktop-open compatibility.

## Example Quality Bar

Every flagship example should include:

- non-trivial input objects
- navigation or TOC
- readable visual hierarchy
- tables and charts
- status/conditional formatting
- metadata or inspection proof
- deterministic output path under `Examples/Documents`
- no dependency on desktop Office for generation
