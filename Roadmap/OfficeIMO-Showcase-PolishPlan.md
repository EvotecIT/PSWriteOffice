# PSWriteOffice Showcase and OfficeIMO Polish Plan

Date: 2026-05-11

This plan compares the current `PSWriteOffice` surface with the newer `OfficeIMO` Word, Excel, and PowerPoint capabilities, then turns the gaps into showcase examples, product polish, and blog deliverables.

The goal is not to prove that the module can add a paragraph or two cells. The goal is to make PSWriteOffice feel like a practical test-drive layer for beautiful, useful, fast Office documents generated from PowerShell.

## Current Position

`PSWriteOffice` is already much stronger than the older gap backlog suggests.

- Word now wraps document creation, readers, paragraphs, lists, tables, conditional table formatting, TOC, bookmarks, fields, footnotes/endnotes, content controls, charts, hyperlinks, document properties, backgrounds, watermarks, protection, mail merge, HTML conversion, and Markdown conversion.
- Excel now wraps sheets, cells, rows, columns, tables, named ranges, formulas, validation, conditional formatting, comments, images and URL images, charts, chart style finishing, pivots, sparklines, TOC/navigation, internal links, URL links, smart hyperlinks, print setup, headers/footers, gridlines, freeze panes, sheet visibility, sorting, autofit, range/read helpers, and workbook summary inspection.
- PowerPoint now wraps slides, titles, text boxes, bullets, notes, sections, tables, images, shapes, charts, backgrounds, layouts, layout placeholders, layout boxes, theme colors/fonts/name, slide transitions, slide sizing, slide import, slide copy, text replacement, inspection helpers, and an initial OfficeIMO designer/deck-plan bridge.

The main next step is not broad primitive coverage. It is opinionated composition: fewer coordinates, more repeatable visual quality, and showcase-grade examples that make the value obvious.

## OfficeIMO Capability Delta

### Word

OfficeIMO.Word offers a deeper document authoring engine than PSWriteOffice currently exposes.

Strongly wrapped now:

- Core document lifecycle and sections
- Headers, footers, page numbers
- Paragraphs, runs through text helpers, lists
- Tables with styles, calculated/projected object data, conditional row formatting, nested table support through table cells
- TOC and field updates
- Content controls: checkbox, date picker, dropdown, combobox, picture control, repeating section
- Hyperlinks and bookmarks
- Footnotes and endnotes
- Document properties
- Backgrounds and watermarks
- Charts
- HTML and Markdown bridges
- Mail merge and text replacement
- Basic read/inspection helpers

Important gaps versus OfficeIMO.Word:

- Text box and richer shape cmdlets in Word
- Cover page helpers
- Append/merge-document helpers
- Image crop, transparency, rotation, wrapping, fixed positioning, and alt text helpers
- Paragraph/run style builder helpers beyond enum styles
- Tab stops, columns, section page setup polish
- Macro inspection/management, probably preview-only
- SmartArt detection/read helpers, not authoring-first
- PDF export through `OfficeIMO.Word.Pdf`, if package scope is approved

Recommended Word focus:

1. Add Word image layout wrappers next because visuals matter for blog-ready reports.
2. Add cover page and document merge helpers for professional report assembly.
3. Add paragraph/run style helpers to reduce repeated formatting in flagship reports.
4. Keep macros, SmartArt mutation, and PDF as explicit scope decisions.

### Excel

OfficeIMO.Excel has become a visually attractive reporting engine with fluent composers, performance policy, rich charts, link helpers, print polish, and report layout helpers.

Strongly wrapped now:

- Workbook create/load/save
- Sheets, cells, rows, columns
- Object-to-table workflows
- Named ranges
- Table of contents and backlinks
- Internal and external hyperlink helpers
- URL images and local images
- Validation and conditional formatting
- Charts plus legend, data label, and style finishing
- Pivot tables and sparklines, with desktop Excel compatibility follow-up required before they return to flagship examples
- Header/footer including images
- Print orientation, margins, page setup, gridlines
- Freeze panes, sort, autofit
- Range/data/table/pivot/used-range readers

Important gaps versus OfficeIMO.Excel:

- Fluent `Compose` / `SheetComposer` report blocks such as title, callout, sections, columns, KPI rows, legends, references, and print defaults
- Column-style-by-header builders for numbers, currency, percentages, dates, durations, and value-based fills
- Execution policy controls and diagnostics for performance-sensitive writes
- Workbook/sheet summary inspection command: shipped as `Get-OfficeExcelSummary`
- More chart finishing: axis titles/text style, axis scale, gridlines, number formats, series fill/line/markers, trendlines, combo/secondary axis, scatter/bubble explicit range helpers
- Pivot table desktop-open compatibility: `Add-OfficeExcelPivotTable -DataFunction` parsing is fixed, but generated PivotTable packages still need an OfficeIMO compatibility pass before they should be used in the flagship workbook.
- Sparkline desktop-open compatibility: generated sparkline packages validate but currently prevent desktop Excel from opening the workbook, so the showcase quarantines them.
- Find/replace wrappers
- Header-aware set/get helpers and editable-row workflows
- A higher-level dashboard/report helper that combines summary sheet, TOC, links, charts, and detail tabs

Recommended Excel focus:

1. Add an `Add-OfficeExcelReportSheet` / `ExcelReportSheet` style composer wrapper that maps to OfficeIMO's fluent report blocks.
2. Add `Set-OfficeExcelColumnStyleByHeader` so financial/status reports become beautiful without range math.
3. Add `Get-OfficeExcelSummary` for test-drive, validation, and blog proof.
4. Add chart-axis/series finishing after the composer layer lands.
5. Expose execution policy only if it stays PowerShell-simple, for example `New-OfficeExcel -ExecutionMode Automatic -MaxDegreeOfParallelism 8`.

### PowerPoint

OfficeIMO.PowerPoint recently moved far ahead with designer/deck-plan capabilities. PSWriteOffice now has an initial bridge for that high-level design system, while deeper diagnostics and remaining semantic slide types still need polish.

Strongly wrapped now:

- Deck create/load/save
- Slides, titles, text boxes, bullets, tables, images, shapes
- Charts including column, pie, doughnut, and scatter
- Background images/colors
- Notes
- Sections
- Theme color/font/name updates
- Layout switching and placeholder text/bounds/style updates
- Layout boxes and columns
- Slide transitions and sizing
- Slide import/copy
- Shape/slide/theme/section/notes inspection
- Initial designer deck rendering through `Add-OfficePowerPointDesignerDeck`
- Semantic deck plans through `New-OfficePowerPointDeckPlan`
- Plan helpers for section, process, card grid, coverage, capability, case study, and logo wall slides

Important gaps versus OfficeIMO.PowerPoint:

- Standalone designer APIs: design brief construction, recipes, recommendations, and deeper alternative inspection
- Remaining semantic deck-plan content: metrics and visual frames
- Composition helpers: title, visual frame, metric strip, callout band, card grid, capability/process/case-study slides
- Layout strategy and content-fit diagnostics
- Shape layout helpers: align, distribute, stack, grid, fit-to-bounds, resize, z-order, duplicate, group/ungroup
- Guides/grid helpers and snap-to-grid
- Rich chart formatting wrappers equivalent to OfficeIMO's title, legend, labels, axis, series, marker, trendline, and scatter axis formatting
- Slide properties such as hidden/show, reorder, and richer duplication controls
- Table cell formatting helpers: merged cells, padding, row heights, borders, autofit, preset styles

Recommended PowerPoint focus:

1. Extend the initial OfficeIMO designer/deck-plan bridge with metrics, visual frames, and richer recommendation diagnostics.
2. Add PowerShell-friendly semantic commands for direct slide composition where the deck-plan layer is too broad, such as `Add-OfficePowerPointMetricSlide` and `Add-OfficePowerPointVisualFrameSlide`.
3. Add shape layout commands for align/distribute/stack/grid/fit because they also improve manual decks.
4. Add richer chart formatting as a follow-up, matching Excel chart finishing concepts where practical.
5. Add validation/diagnostic output so a script can say why a deck layout was selected and whether content fit cleanly.

## Showcase Examples

The examples should be complex enough to act as product demos and regression smoke tests.

### Word Showcase: `Examples/Showcase/Showcase-Word-ExecutiveReport.ps1`

Scenario: an executive service-health report generated from PowerShell objects.

Must demonstrate:

- cover/title page or strong opening section
- header/footer, page numbers, document properties
- TOC and multiple heading levels
- executive summary with callouts
- numbered and bulleted lists
- status table with conditional formatting
- chart section
- hyperlink/bookmark navigation
- content controls for approvals
- watermark or background
- image/logo once image layout wrappers are strong enough
- optional HTML or Markdown export sidecar

Needed polish before or during:

- footnote/endnote wrappers: shipped
- better image layout/crop/alt-text wrappers
- cover page or opening-section helper

### Excel Showcase: `Examples/Showcase/Showcase-Excel-OperationalDashboard.ps1`

Scenario: a multi-sheet operational workbook with summary, details, trend, and appendix tabs.

Must demonstrate:

- summary dashboard sheet
- TOC with links and backlinks
- KPI blocks
- status legend
- object tables with friendly headers
- conditional color scale/data bars/icon sets
- validation list for status
- owner summary sheet while PivotTable compatibility is repaired
- sparklines once desktop-open compatibility is repaired
- charts with styled labels/legend
- internal links by header
- smart external links
- URL/local image
- print setup and header/footer logo
- hidden notes/config sheet
- `Get-OfficeExcelSummary` validation

Needed polish before or during:

- report composer wrapper over OfficeIMO `SheetComposer`
- column-style-by-header wrapper
- workbook summary inspection
- chart axis/series finishing wrappers

### PowerPoint Showcase: `Examples/Showcase/Showcase-PowerPoint-ServiceBrief.ps1`

Scenario: a service/consulting deck with executive story, process, KPIs, coverage, proof points, charts, and appendix.

Must demonstrate:

- branded theme and fonts
- slide size and transitions
- sections
- title/section slide
- metric strip
- process slide
- card grid slide
- coverage/map-like slide
- chart slide
- table slide
- image/background slide
- speaker notes
- imported/copy slide appendix
- inspection summary after save

Needed polish before or during:

- PowerPoint designer/deck-plan wrappers: initial bridge shipped
- semantic slide commands: initial plan helpers shipped
- shape layout commands
- chart formatting wrappers
- table cell formatting wrappers

## Blog Series

Create three separate posts in `C:\Support\GitHub\Website.Contributions` once the showcase examples generate real artifacts.

### Post 1: Word

Working title: `Build a polished Word executive report from PowerShell`

Assets:

- generated DOCX
- cover/hero image
- screenshot of cover page / TOC / chart page
- code excerpt from `Showcase-Word-ExecutiveReport.ps1`

Story:

- why Word output is useful for operational reporting
- how sections, TOC, tables, charts, approvals, and document metadata fit together
- what OfficeIMO owns versus what PSWriteOffice makes PowerShell-native

### Post 2: Excel

Working title: `Build an Excel operational dashboard from PowerShell`

Assets:

- generated XLSX
- cover/hero image
- screenshots of summary dashboard, detail table, chart/pivot page
- code excerpt from `Showcase-Excel-OperationalDashboard.ps1`

Story:

- not just exporting rows: navigation, KPIs, formatting, validation, pivots, charts, print setup
- how the workbook stays useful to a human after generation
- where performance and deterministic Open XML output matter

### Post 3: PowerPoint

Working title: `Build a beautiful PowerPoint service brief from PowerShell`

Assets:

- generated PPTX
- cover/hero image
- screenshots of title slide, process slide, metric slide, chart slide
- code excerpt from `Showcase-PowerPoint-ServiceBrief.ps1`

Story:

- from PowerShell objects to editable business deck
- why semantic deck helpers are more useful than coordinates
- design recommendations, slide plans, notes, and editable output

## Visual Pipeline

Blog posts should include real visuals of generated output, not only decorative covers.

Preferred pipeline:

1. Generate DOCX/XLSX/PPTX examples into `Examples/Documents`.
2. Export representative pages/slides/sheets to PNG/WebP.
3. Use image model v2 only for covers or explanatory infographics, not as a substitute for output screenshots.
4. Store article images under each post's `images` folder.
5. Keep the generated Office artifacts linked from the article or release assets if the blog platform allows downloads.

Possible export paths:

- PowerPoint: export slides through installed Office PowerPoint if available; otherwise evaluate LibreOffice/headless fallback.
- Word: export to PDF/PNG through Word if available; once `OfficeIMO.Word.Pdf` is in scope, prefer pure library export for repeatable previews.
- Excel: export selected sheets through Excel if available; otherwise add an HTML/PNG preview helper later.

## Implementation Order

### Slice 1: Update the truth

- Refresh `Roadmap/OfficeIMO-GapBacklog.md` so completed items are not treated as missing.
- Add this showcase/polish plan.
- Add a support matrix page with `wrapped now`, `needs PSWriteOffice wrapper`, `needs OfficeIMO engine work`, and `intentionally deferred`.

### Slice 2: Excel-first showcase

Excel is closest to showcase-ready today.

Status: initial implementation completed. `Showcase-Excel-OperationalDashboard.ps1` generates a six-sheet workbook with tables, charts, formulas, owner summary, validation, conditional formatting, navigation, and hidden notes. The pivot `DataFunction` parsing issue is fixed, and PivotTable/sparkline desktop-open compatibility is now tracked separately instead of being hidden inside the flagship example.

- Build `Showcase-Excel-OperationalDashboard.ps1` using existing cmdlets.
- Add only the smallest missing wrapper needed for visual quality.
- Fix the `Add-OfficeExcelPivotTable -DataFunction` parsing bug found by the showcase.
- Validate generated workbook structure and chart XML.
- Draft the Excel blog post once screenshots exist.

### Slice 3: PowerPoint designer bridge

PowerPoint has the largest value gap because OfficeIMO's designer layer is new and visually attractive.

Status: initial implementation completed. `Showcase-PowerPoint-ServiceBrief.ps1` generates an eight-slide service brief with designer-composed plan slides plus chart/table follow-up slides, speaker notes, sections, and transitions. A focused PowerPoint test now locks the deck-plan DSL and designer bridge.

- Add design brief/deck-plan wrappers in PSWriteOffice.
- Build `Showcase-PowerPoint-ServiceBrief.ps1` on semantic slide commands.
- Add tests that validate slide count, sections, notes, charts, and no Open XML repair prompts.
- Draft the PowerPoint blog post with real slide screenshots.

### Slice 4: Word report polish

Word is mature, but the showcase needs richer report polish.

Status: initial implementation completed. `Showcase-Word-ExecutiveReport.ps1` generates an executive report with header/footer, generated banner image, TOC, conditional scorecard table, chart, action table, bookmark/hyperlink navigation, approval controls, watermark, document properties, footnote, and endnote. Footnote/endnote add/read wrappers and a parameter-set fix for TOC updates landed with tests.

- Add footnote/endnote wrappers.
- Add image layout/alt-text wrappers.
- Consider cover page helper if OfficeIMO exposes a stable template API.
- Build `Showcase-Word-ExecutiveReport.ps1`.
- Draft the Word blog post with real page screenshots.

## Acceptance Criteria

- Each product has one flagship showcase script that creates a non-trivial artifact from PowerShell objects.
- Each flagship artifact includes navigation, visual hierarchy, structured data, and at least one chart or visual.
- Each example is fast enough to run as a smoke test and deterministic enough to inspect in CI.
- Each blog post includes runnable code plus real screenshots of generated output.
- Decorative generated covers are clearly separate from screenshots of actual output.
- Missing capability is fixed at the lowest sensible layer: OfficeIMO for engine behavior, PSWriteOffice for PowerShell ergonomics.
