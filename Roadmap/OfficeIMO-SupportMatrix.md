# PSWriteOffice and OfficeIMO Support Matrix

Date: 2026-05-11

This matrix is a quick planning companion for `OfficeIMO-Showcase-PolishPlan.md`.

## Legend

- `Wrapped`: available through PSWriteOffice today.
- `Partial wrapper`: a useful PSWriteOffice surface exists, but deeper OfficeIMO capability remains to expose.
- `Wrapper gap`: OfficeIMO appears to have a useful capability that PSWriteOffice should expose.
- `Engine gap`: needs OfficeIMO work first or needs upstream API stabilization.
- `Deferred`: useful, but should not be pulled into core scope without an explicit decision.

## Word

| Capability | Status | Notes |
| --- | --- | --- |
| Create/load/save documents | Wrapped | `New-OfficeWord`, `Get-OfficeWord`, `Save-OfficeWord`, `Close-OfficeWord` |
| Sections, headers, footers, page numbers | Wrapped | Good enough for report examples |
| Paragraphs, text, lists | Wrapped | Needs richer run/style builder later |
| Tables from objects | Wrapped | Includes conditional formatting and table-cell helpers |
| TOC and field updates | Wrapped | `Add/Set/Get/Update/Remove-OfficeWordTableOfContent`, `Update-OfficeWordFields` |
| Bookmarks and hyperlinks | Wrapped | External and anchor links are exposed |
| Document properties | Wrapped | Built-in/custom property surface exists |
| Backgrounds and watermarks | Wrapped | Color/image backgrounds and watermarking exist |
| Content controls | Wrapped | Checkbox, date, dropdown, combo, picture, repeating section |
| Charts | Wrapped | Good enough for showcase reports |
| HTML and Markdown conversion | Wrapped | Useful for sidecar previews/blog code |
| Mail merge | Wrapped | Suitable for practical examples |
| Footnotes/endnotes | Wrapped | Add/read wrappers return document-safe note snapshots |
| Advanced image layout | Wrapper gap | Crop, transparency, rotation, wrapping, fixed positioning, alt text |
| Text boxes and shapes | Wrapper gap | OfficeIMO.Word supports richer shape scenarios than PSWriteOffice exposes |
| Cover pages | Wrapper gap | Prefer template-driven wrapper if stable |
| Append/merge documents | Wrapper gap | Useful for report packs and appendices |
| Paragraph/run style builders | Wrapper gap | Useful once flagship examples need less raw styling |
| Macros | Deferred | Keep preview-only if added |
| SmartArt authoring | Deferred | Detection/read helpers are safer first |
| PDF export | Deferred | Requires package-scope approval for `OfficeIMO.Word.Pdf` |

## Excel

| Capability | Status | Notes |
| --- | --- | --- |
| Create/load/save workbook | Wrapped | `New/Get/Save/Close-OfficeExcel` |
| Sheets, cells, rows, columns | Wrapped | Strong primitive coverage |
| Tables from objects | Wrapped | Core reporting path works |
| Named ranges | Wrapped | Includes hidden and validation-mode support |
| TOC and workbook navigation | Wrapped | Includes internal links/backlinks |
| Range/data/table/pivot readers | Wrapped | Good read-back foundation |
| Validation | Wrapped | List, whole number, decimal, date, time, text length, custom formula |
| Conditional formatting | Wrapped | Rules, color scale, data bar, icon set |
| Charts | Wrapped | Add chart plus style, legend, data labels |
| Pivot tables and sparklines | Partial wrapper | Cmdlets exist, but desktop Excel open compatibility needs engine follow-up before flagship examples should rely on them |
| Images and URL images | Wrapped | Includes in-sheet and header/footer image paths |
| Hyperlinks | Wrapped | Raw, smart, host, internal, URL by header |
| Print setup | Wrapped | Orientation, margins, page setup, gridlines, freeze |
| Sort/autofit | Wrapped | Direct helpers exist |
| Fluent report composer | Wrapper gap | OfficeIMO `Compose` / `SheetComposer` should become a PowerShell report sheet DSL |
| KPI, legend, callout, columns blocks | Wrapper gap | Best route to beautiful default dashboards |
| Column style by header | Wrapper gap | Currency/percent/date/status formatting without range math |
| Execution policy/diagnostics | Wrapper gap | Expose only if simple and useful from PowerShell |
| Workbook summary inspection | Wrapped | `Get-OfficeExcelSummary` reports workbook shape, charts, tables, pivots, sparklines, links, comments, named ranges, sheet visibility, and used ranges |
| Rich chart axis/series formatting | Wrapper gap | Axis titles, gridlines, scale, number format, trendlines, markers, combo/secondary axis |
| Find/replace and editable rows | Wrapper gap | Useful for maintenance workflows |
| Google Sheets bridge | Deferred | Explicit package-scope expansion only |

## PowerPoint

| Capability | Status | Notes |
| --- | --- | --- |
| Create/load/save decks | Wrapped | `New/Get/Save-OfficePowerPoint` |
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
| Designer brief/recipes/directions | Partial wrapper | `Add-OfficePowerPointDesignerDeck` exposes accent, seed, purpose, creative direction, layout strategy, alternatives, theme application, preview, and summaries |
| Deck plan and semantic slides | Partial wrapper | `New-OfficePowerPointDeckPlan` plus section, process, card grid, coverage, capability, case study, and logo wall helpers; metrics and visual frames remain |
| Content-fit diagnostics | Partial wrapper | Preview/render summaries exist; richer fit warnings should still be exposed |
| Shape layout helpers | Wrapper gap | Align, distribute, stack, grid, fit, resize, z-order, duplicate, group |
| Guides and snap-to-grid | Wrapper gap | Useful for manual layout scripts |
| Rich chart formatting | Wrapper gap | Match OfficeIMO chart title/legend/axis/series/marker/trendline APIs |
| Table cell formatting | Wrapper gap | Merged cells, padding, borders, row heights, autofit, preset style wrappers |
| Slide hidden/reorder controls | Wrapper gap | Useful deck maintenance commands |

## Recommended First PRs

1. Blog draft preparation once screenshots are produced.
2. Excel report composer wrapper after the current operational dashboard showcase.
3. PowerPoint metrics/visual-frame helpers, fit diagnostics, and shape layout polish after the current designer bridge.
4. Word image layout, cover-page, append/merge, and richer style helpers after the current executive report showcase.

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
