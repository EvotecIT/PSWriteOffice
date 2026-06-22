# PSWriteOffice and OfficeIMO Support Matrix

Date: 2026-06-21

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

## Latest OfficeIMO Release Review

Latest checked release: `OfficeIMO 2026.06.16-19.47.03`
(`OfficeIMO-v20260616194703`).

The newest OfficeIMO releases since the previous matrix are mostly engine
hardening and fidelity improvements rather than brand-new PSWriteOffice command
families:

- Security/resource guardrails: chart number-format parsing, VML image reads,
  Visio comment XML loading, Excel style table preallocation, formula range
  materialization, and HTML table span expansion now have tighter bounds in
  OfficeIMO.
- PDF fidelity: document table workflows, typography/conversion diagnostics,
  HTML PDF proof contracts, PDF page chunk action deduplication, and default
  table-style preservation in converters improved in OfficeIMO.Pdf and related
  converter packages.
- Reader/ingestion: YAML, RTF, EPUB, ZIP, HTML, CSV, JSON, XML, Visio, and PDF
  modular adapters are now part of the package set PSWriteOffice references.
  PSWriteOffice already registers these adapters through the Reader command
  utilities, so the next gap is examples/profiles rather than package wiring.
- Word/Markdown/PDF: Word Markdown conversion fidelity, Word PDF conversion
  fidelity, native Markdown projection, and Word block insertion ordering
  improved in OfficeIMO. Existing PSWriteOffice conversion commands benefit from
  the package bump through the normal OfficeIMO APIs.
- Excel: direct tabular range handling improved in OfficeIMO.Excel. Existing
  import/export/range/table commands should stay thin and use the normal
  OfficeIMO reader/writer APIs.

Exposure guidance:

- Do not add PSWriteOffice-only fast paths for save/performance. Keep using the
  standard OfficeIMO APIs and let OfficeIMO choose optimized internals.
- Add focused Reader profiles next: folder ingest presets, deterministic JSONL
  export, format include/exclude presets, and cookbook examples that prove
  YAML/RTF/EPUB/ZIP ingestion in real workflows.
- Add PDF table style depth next because OfficeIMO now preserves default table
  styles through converters and already owns rich `PdfTableStyle` behavior.
- Improve examples for Word Markdown/PDF and HTML/PDF conversion so users can
  see the fidelity improvements without needing to read OfficeIMO release notes.
- Keep Google Workspace, Markup, and MarkdownRenderer packages out of the module
  until PSWriteOffice gets an intentional command surface for those workflows.

## Latest OfficeIMO PR Assessment

Checked on 2026-06-21 against the latest merged OfficeIMO PRs:

- `#1984 Improve Markdown PDF visual rendering`
- `#1983 Add shared HTML engine platform contracts`
- `#1982 Improve Excel and PowerPoint PDF conversion fidelity`
- `#1981 Improve native Word PDF engine parity`
- `#1980 Add shared HTML diagnostics and gallery contracts`

The Word-specific follow-up also checked `#1978 Improve Word market readiness`
because it is the recent PR that introduced the structured Word comparison
result shape.

### PR 1984: Markdown PDF visuals

OfficeIMO added richer Markdown-to-PDF rendering: shared Markdown visual themes,
PDF-specific figure styling, chart semantic-fence rendering, image-only
paragraph handling, captions/placeholders, front-matter theme resolution, and
conversion warnings/reports.

PSWriteOffice currently exposes `New-OfficeMarkdown -PdfPath` and
`Save-OfficeMarkdown -PdfPath`, but both call `SaveAsPdf(path)` without
friendly access to `MarkdownPdfSaveOptions`.

Recommended PSWriteOffice work:

- Add a shared Markdown PDF option builder used by both `New-OfficeMarkdown` and
  `Save-OfficeMarkdown`.
- Expose a compact parameter surface first: `-PdfTheme`, `-PdfOptions`,
  `-PdfFontFamily`, `-PdfBaseDirectory`, `-PdfIncludeLocalImages`,
  `-PdfIncludeDataUriImages`, `-PdfDefaultImageWidth`,
  `-PdfDefaultImageHeight`, `-PdfFrontMatterRenderMode`,
  `-PdfUseFrontMatterVisualTheme`, `-PdfUseFrontMatterMetadata`,
  `-PdfCreateOutlineFromHeadings`, and an optional warning/report output hook.
- Add examples proving Markdown chart fences, image captions, front matter, and
  the built-in `Plain`, `WordLike`, `TechnicalDocument`, `GitHubLike`,
  `Compact`, and `Report` visual themes.

### PR 1983: Shared HTML engine platform

OfficeIMO added shared HTML conversion profiles, logical document/resource
manifests, computed-style support, diagnostic catalogs, round-trip scoring, and
profile contracts. Existing PSWriteOffice HTML commands already expose basic
HTML-to-Word, HTML-to-Markdown, and HTML-to-PDF workflows, but not the new shared
profile and diagnostics model.

Recommended PSWriteOffice work:

- Add `-ConversionProfile` to HTML-backed commands where it maps cleanly:
  `ConvertFrom-OfficeWordHtml`, `ConvertFrom-OfficeMarkdownHtml`, and
  `ConvertFrom-OfficePdfHtml`. For PDF, keep `HtmlPdfProfile` as the rendering
  path selector and use `HtmlConversionProfile` for semantic/document/print
  intent when building nested options.
- Add diagnostics/report output hooks for HTML conversions, such as
  `-DiagnosticVariable`, `-ConversionReportVariable`, or a `-PassReport` mode.
- Add safety/profile switches for common workflows: untrusted HTML, trusted
  document HTML, and visual/print review. Do not expose every low-level URL
  policy collection as first-class parameters initially; accept `-Options` for
  advanced callers.
- Add cookbook examples for invoice/report/contract/email/dashboard-print
  scenarios that show blocked resources and diagnostics.

### PR 1982: Excel and PowerPoint PDF fidelity

OfficeIMO improved Excel chart title propagation, Excel `FitToHeight` scaling,
PowerPoint common preset geometry rendering, inherited-layout shapes, and
straight connector rendering in PDF output.

PSWriteOffice already exposes the relevant inputs:

- `Set-OfficeExcelPageSetup -FitToHeight`
- `Set-OfficeExcelPrintArea`
- Excel chart commands and PDF sidecar export through `New/Save-OfficeExcel`
- `Add-OfficePowerPointShape` with Open XML shape preset names
- PowerPoint PDF sidecar export through `New/Save-OfficePowerPoint`

Recommended PSWriteOffice work:

- No new wrapper parameters are needed for this PR.
- Add or refresh examples that deliberately prove Excel chart titles,
  `FitToHeight`, print areas, PowerPoint triangles/arrows/connectors, and
  layout-inherited shapes survive `-PdfPath`.
- Add focused Pester/readback smoke tests only for the PSWriteOffice command
  contract. The shape geometry and layout fidelity remain OfficeIMO engine
  tests.

### Word diffs from PRs 1981, 1980, 1978

OfficeIMO's recent Word work adds meaningful wrapper opportunities:

- `#1981` improved native Word PDF parity across headers/footers, lists,
  columns, paragraph spacing, tables, structured blocks, table style defaults,
  conditional borders/padding, and diagnostics. PSWriteOffice currently exposes
  only `-PdfPath`; it should add a shared Word PDF option builder for
  `New-OfficeWord` and `Save-OfficeWord` that maps to
  `OfficeIMO.Word.Pdf.PdfSaveOptions`.
- `#1980` added shared HTML diagnostics and gallery contracts. Together with
  `#1983`, this makes HTML import/reporting a first-class surface rather than
  just a conversion helper.
- `#1978` added `WordDocumentComparer.CompareStructure()` with deterministic
  `WordComparisonResult` and findings across paragraphs, tables, rows, cells,
  images, and block order. PSWriteOffice has no Word comparison cmdlet today.

Recommended PSWriteOffice work:

- Add `Compare-OfficeWordDocument` with two modes: a revision-mark document
  mode using `WordDocumentComparer.Compare()` and a machine-readable mode using
  `CompareStructure()`. Let callers save the diff document with `-OutputPath`
  and optionally return findings through `-PassThru` or `-AsJson`.
- Add Word PDF options to `New-OfficeWord` and `Save-OfficeWord`: `-PdfOptions`,
  `-PdfFontFamily`, `-PdfAllowSystemFontEmbedding`, `-PdfPageSize`,
  `-PdfOrientation`, `-PdfTitle`, `-PdfAuthor`, `-PdfSubject`, `-PdfKeywords`,
  `-PdfIncludePageNumbers`, `-PdfPageNumberFormat`, `-PdfDefaultTableBorders`,
  and a conversion report/warning output hook.
- Expand `ConvertFrom-OfficeWordHtml` with `-ConversionProfile`,
  `-HtmlProfile` presets for OfficeIMO, untrusted HTML, and trusted document
  HTML, plus diagnostics/report output. Keep an advanced `-Options` escape
  hatch for low-level URL/resource policies.
- Expand `ConvertFrom-OfficeWordMarkdown` with `-Theme`,
  `-AllowDataUriImages`, `-MaxDataUriImageBytes`, and warning/image-layout
  diagnostic output. The command already exposes the core local/remote image
  and fitting knobs.
- Add examples that prove the new OfficeIMO engine work through PSWriteOffice:
  Word PDF with headers, lists, columns, table styles, and structured blocks;
  Word HTML import with diagnostics; Word Markdown import with themed output;
  and a Word compare report that can be used in CI.

## Markdown and PDF Exposure Assessment

OfficeIMO now has enough shared Markdown/PDF option objects that PSWriteOffice
should expose workflow-sized presets rather than one switch for every engine
property.

### Markdown engine capabilities worth exposing

OfficeIMO.Markdown provides:

- Reader dialect profiles: `OfficeIMO`, `CommonMark`,
  `GitHubFlavoredMarkdown`, and `Portable`.
- Reader safety/compatibility switches for front matter, callouts, task lists,
  tables, definition lists, TOC placeholders, footnotes, raw/inline HTML,
  autolinks, URL-scheme restrictions, max input length, and base URI.
- Markdown writer profiles: OfficeIMO, portable, and HTML-image output, plus
  image rendering mode, line endings, and unordered list marker.
- Shared visual themes: `Plain`, `WordLike`, `TechnicalDocument`,
  `GitHubLike`, `Compact`, and `Report`, with optional color schemes and table
  styling.
- Input-normalization presets for loose documentation imports and
  transcript/model-output cleanup.
- Semantic visual transforms that upgrade chart/network/dataview JSON fences
  into typed semantic fenced blocks.
- HTML conversion controls for base URI, script/style stripping, unsupported
  HTML preservation, base64 image handling, listing-card metadata suppression,
  markdown write options, and max input length.

Recommended PSWriteOffice work:

- Implemented first tranche: shared reader/write/PDF option helpers now expose
  friendly Markdown profiles, URL safety, input normalization, HTML visual
  themes, Markdown writer profiles, image rendering mode, line endings,
  unordered list markers, and Markdown PDF metadata/theme/image/report options
  across the existing Markdown commands, including `Get-OfficeMarkdown`.
- Add visual-fence support as one focused switch, for example
  `-EnableVisualFences`, with a small enum for preserve/generic/IntelligenceX
  fence language output. Keep custom parser/renderer extensions as `-Options`.
- Keep expanding examples so parsing, HTML, Word, and PDF flows show the same
  dialect and safety surface.

Defer:

- Do not expose parser extension collections, inline renderer delegates,
  custom HTML block converters, or visual round-trip hint collections as named
  PowerShell parameters. Accept the native `-Options` object for those advanced
  cases.

### Markdown-to-PDF capabilities worth exposing

OfficeIMO.Markdown.Pdf provides `MarkdownPdfSaveOptions` with:

- PDF engine pass-through through `PdfOptions`.
- Shared `MarkdownVisualTheme` and PDF-specific `MarkdownPdfVisualTheme`.
- Built-in PDF themes: `Plain`, `WordLike`, `TechnicalDocument`,
  `GitHubLike`, `Compact`, and `Report`.
- Front matter theme/metadata handling, front matter render mode, first heading
  as title, and outline creation from headings.
- Explicit PDF metadata: title, author, subject, and keywords.
- Local/data URI/remote image controls, base directory confinement, max image
  byte limits, fallback image dimensions, warnings, and conversion report.

Recommended PSWriteOffice work:

- Implemented first tranche: `New-OfficeMarkdown` and `Save-OfficeMarkdown`
  share a Markdown PDF option builder and expose first-class parameters for
  `-PdfTheme`, `-PdfOptions`, `-PdfFontFamily`, metadata, base directory, image
  controls, front matter behavior, first-heading title, heading outlines,
  warning capture, and conversion-report capture.
- Consider a direct `ConvertFrom-OfficeMarkdownPdf` cmdlet for file/text/document
  input parity.

### PDF engine capabilities worth exposing

PSWriteOffice already wraps a substantial part of OfficeIMO.Pdf: document
creation, themes, metadata, page setup, headers/footers, background color/image,
background shapes, page borders, text/heading/list/table/panel/image/row
composition, rich text runs, bookmarks, attachments, form fields, stamps, page
operations, readback, preflight, compliance readiness, and HTML/PDF conversion.

The best remaining PDF gaps are not basic block authoring; they are workflow
options that make generated PDFs easier to ship:

- Add `-Theme` directly to `New-OfficePdf` in addition to the DSL
  `PdfTheme` command, and include a `Plain` preset if OfficeIMO adds or exposes
  one for PDF themes.
- Add catalog/viewer options: catalog page mode/layout, open action,
  display-document-title, page labels, outline expansion level, document
  language, and URI base.
- Add diagnostics/report output for generated PDFs, HTML-to-PDF, Word-to-PDF,
  and Markdown-to-PDF using the shared `PdfConversionReport` pattern.
- Expand compliance/invoice workflows: expose Factur-X/ZUGFeRD groundwork,
  invoice XML attachment, output intent selection, and embedded-file
  relationships as a focused invoice/compliance command rather than making users
  build raw `PdfOptions`.
- Add generated-PDF font fallback helpers for common host font families, but
  keep arbitrary embedded-font collection editing as an advanced `PdfOptions`
  scenario.
- Add generated annotation helpers for text/free-text/highlight annotations if
  OfficeIMO.Pdf block APIs are stable enough for thin cmdlets.

Defer:

- Do not mirror every `PdfOptions` catalog/font/style property as a top-level
  `New-OfficePdf` parameter. Keep advanced composition through DSL commands,
  hashtable style builders, or native `PdfOptions` objects.

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
| HTML and Markdown conversion | Partial wrapper | Useful for sidecar previews/blog code; latest OfficeIMO profile/theme/diagnostic options need friendlier parameters |
| Mail merge | Wrapped | Suitable for practical examples |
| Footnotes/endnotes | Wrapped | Add/read wrappers return document-safe note snapshots |
| Page setup and columns | Wrapped | `Set-OfficeWordPageSetup` covers page size, orientation, margins, and columns |
| Advanced image layout | Partial wrapper | `Get/Set-OfficeWordImage` exposes crop, rotation, flip, wrapping, metadata, and visibility; fixed-position semantics remain engine-led |
| Text boxes and shapes | Partial wrapper | `Add/Get/Set-OfficeWordShape` exposes basic shape authoring and styling; text boxes and richer templates remain |
| Cover pages | Wrapped | `Add-OfficeWordCoverPage` exposes stable OfficeIMO templates and basic cover metadata |
| Append/merge documents | Wrapped | `Join-OfficeWordDocument` appends one or more documents into a base document |
| Document comparison | Wrapper gap | OfficeIMO exposes revision-mark comparison and structured findings through `WordDocumentComparer`; add `Compare-OfficeWordDocument` |
| Equations and tab stops | Wrapped | `Add-OfficeWordEquation` and `Add-OfficeWordTabStop` expose stable OfficeIMO.Word APIs |
| Document statistics | Wrapped | `Get-OfficeWordStatistics` exposes page/paragraph/word/object counts |
| Macros | Deferred | Keep preview-only if added |
| SmartArt authoring | Deferred | Detection/read helpers are safer first |
| PDF export | Partial wrapper | `New/Save-OfficeWord -PdfPath` use OfficeIMO.Word.Pdf sidecar export; expose `PdfSaveOptions` and report/warning hooks next |

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

1. Word comparison, Word PDF options/report hooks, and HTML/Markdown conversion diagnostics.
2. Word run/paragraph style, row/column table mutation, and text box helpers.
3. PowerPoint metrics/visual-frame helpers, fit diagnostics, and shape layout polish.
4. Visio grouping/layer/layout and semantic diagram builder wrappers after OfficeIMO.Visio stabilizes the high-level workflows.
5. OfficeIMO engine confidence for Excel pivot/sparkline desktop-open compatibility.

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
