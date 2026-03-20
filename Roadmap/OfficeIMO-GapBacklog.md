# PSWriteOffice OfficeIMO Gap Backlog

This backlog captures the main feature gaps between `PSWriteOffice` and the broader sibling `OfficeIMO.*` ecosystem as of 2026-03-20.

It is intentionally opinionated:

- Keep `PSWriteOffice` thin and PowerShell-native.
- Add feature logic to `OfficeIMO` first when the C# API is not ready yet.
- Prioritize wrappers that unlock a lot of user value without pulling in whole new package families.

## Priority bands

- `P1`: high-value, in-scope additions that fit the current module direction
- `P2`: useful next additions after the `P1` surface is stable
- `P3`: explicit scope expansion into additional `OfficeIMO.*` packages

## Recommended delivery order

1. Markdown advanced blocks and reader options
2. Word hyperlinks, document properties, backgrounds, and mail merge
3. PowerPoint charts, backgrounds, and layout-box helpers
4. CSV schema validation and streaming helpers
5. Additional Word/PowerPoint rich-content helpers
6. Explicit package-scope expansion

## Word

### P1

- [ ] Add hyperlink cmdlets on top of `OfficeIMO.Word.WordHyperLink`
  - Proposed surface: `Add-OfficeWordHyperlink`, `Get-OfficeWordHyperlink`
  - Support external URLs and bookmark anchors
  - Add DSL alias if it reads cleanly inside `WordParagraph`

- [ ] Add built-in and custom document property cmdlets
  - Proposed surface: `Get-OfficeWordDocumentProperty`, `Set-OfficeWordDocumentProperty`
  - Cover common built-ins first: title, subject, creator, keywords, description
  - Support custom properties once naming and typing rules are clear

- [ ] Add document background cmdlets
  - Proposed surface: `Set-OfficeWordBackground`
  - Support solid color and background image modes

- [ ] Add mail-merge wrapper cmdlets
  - Proposed surface: `Invoke-OfficeWordMailMerge`
  - Start with hashtable/object-driven field replacement
  - Keep advanced field-preservation switches for a second pass if needed

### P2

- [ ] Add footnote read/write cmdlets
  - Proposed surface: `Add-OfficeWordFootnote`, `Get-OfficeWordFootnote`

- [ ] Add text box and shape cmdlets
  - Proposed surface: `Add-OfficeWordTextBox`, `Add-OfficeWordShape`, `Get-OfficeWordTextBox`

- [ ] Add equation wrappers
  - Proposed surface: `Add-OfficeWordEquation`, `Get-OfficeWordEquation`

- [ ] Add SmartArt wrappers only after a minimal, predictable PowerShell shape is agreed

- [ ] Add cover page helpers if they can stay thin and template-driven

### P3

- [ ] Add macro inspection/management cmdlets
  - Proposed surface: `Get-OfficeWordMacro`, `Remove-OfficeWordMacro`, `Save-OfficeWordMacro`
  - Keep this behind an explicit preview marker if needed

- [ ] Add embedded object helpers
  - Proposed surface: `Add-OfficeWordEmbeddedObject`

## Excel

The core Excel surface is already comparatively strong. The next backlog here is more about inspection and reporting ergonomics than large missing package coverage.

### P1

- [ ] Add workbook/sheet summary inspection cmdlet
  - Proposed surface: `Get-OfficeExcelSummary`
  - Include sheets, tables, named ranges, pivot tables, used range, hyperlinks, and visibility

- [ ] Add more object-first inspection helpers where upstream OfficeIMO APIs are already stable
  - Focus on workbook navigation and reporting metadata, not DSL cleverness

### P2

- [ ] Review whether chart axis/series formatting should be wrapped next
  - Only proceed if the upstream OfficeIMO Excel API is stable enough to keep cmdlets thin

- [ ] Add higher-level dashboard/reporting helpers on top of TOC, internal links, and URL-link primitives

### P3

- [ ] Evaluate `OfficeIMO.Excel.GoogleSheets` only as an explicit scope expansion
  - Do not mix this into the core Excel backlog until package scope is approved

## PowerPoint

### P1

- [ ] Add PowerPoint chart cmdlets
  - Proposed surface: `Add-OfficePowerPointChart`
  - Start with a small set of common chart types and simple data binding
  - Follow with thin formatting cmdlets only if the API shape is clean

- [ ] Add slide background cmdlets
  - Proposed surface: `Set-OfficePowerPointBackground`
  - Support color and image backgrounds

- [ ] Add layout-box helpers for script authors
  - Proposed surface: `Get-OfficePowerPointLayoutBox`
  - Include content box and multi-column helpers backed by slide size

### P2

- [ ] Add image update helpers if they can be expressed cleanly from PowerShell
  - Proposed surface: `Update-OfficePowerPointImage`

- [ ] Add notes master helpers if upstream behavior is stable enough

- [ ] Add richer chart formatting once the first chart cmdlet is proven
  - Titles, legends, labels, markers, and series formatting should be separate follow-up cmdlets

## Markdown

This is the cleanest high-payoff gap area because the upstream `OfficeIMO.Markdown` library already has a much broader model than `PSWriteOffice` exposes today.

### P1

- [ ] Add front matter support
  - Proposed surface: `Add-OfficeMarkdownFrontMatter`

- [ ] Add TOC helpers
  - Proposed surface: `Add-OfficeMarkdownTableOfContents`
  - Support top-level and section-scoped variants

- [ ] Add task-list support
  - Proposed surface: `Add-OfficeMarkdownTaskList`

- [ ] Add footnote support
  - Proposed surface: `Add-OfficeMarkdownFootnote`

- [ ] Add definition-list support
  - Proposed surface: `Add-OfficeMarkdownDefinitionList`

- [ ] Add details/summary blocks
  - Proposed surface: `Add-OfficeMarkdownDetails`

### P2

- [ ] Add semantic fenced block support
  - Proposed surface: `Add-OfficeMarkdownSemanticFence`
  - Keep this generic rather than hard-coding chart-only behavior

- [ ] Add raw HTML block support
  - Proposed surface: `Add-OfficeMarkdownHtml`

- [ ] Add reader-option cmdlets or parameters for advanced parse features
  - Front matter, task lists, footnotes, definition lists

## CSV

### P1

- [ ] Add schema-validation support
  - Proposed surface: `Test-OfficeCsvSchema` or `Assert-OfficeCsvSchema`
  - Start with validation reporting before adding mutation helpers

- [ ] Add streaming mode support
  - Proposed surface: `Get-OfficeCsvData -Mode Stream`
  - Validate end-to-end pipeline behavior and disposal semantics

- [ ] Add object mapping helpers
  - Proposed surface: `ConvertFrom-OfficeCsv`
  - Map rows to typed objects or PowerShell objects using column rules

### P2

- [ ] Add row/column transform helpers if they stay composable
  - Filtering, sorting, adding/removing columns, materialization

## Explicit Scope Expansion

These are real `OfficeIMO.*` gaps, but they should be treated as package-scope decisions rather than normal backlog items.

### P3

- [ ] `OfficeIMO.Word.Pdf`
- [ ] `OfficeIMO.Pdf`
- [ ] `OfficeIMO.Visio`
- [ ] `OfficeIMO.Reader`
- [ ] `OfficeIMO.Reader.Zip`
- [ ] `OfficeIMO.Reader.Epub`
- [ ] `OfficeIMO.Reader.Text`
- [ ] `OfficeIMO.Reader.Html`
- [ ] `OfficeIMO.Reader.Csv`
- [ ] `OfficeIMO.Reader.Json`
- [ ] `OfficeIMO.Reader.Xml`
- [ ] `OfficeIMO.Epub`
- [ ] `OfficeIMO.Markdown.Html`
- [ ] `OfficeIMO.MarkdownRenderer`
- [ ] `OfficeIMO.GoogleWorkspace`
- [ ] `OfficeIMO.Word.GoogleDocs`
- [ ] `OfficeIMO.Excel.GoogleSheets`

## Cross-cutting follow-up

- [ ] Add a support matrix doc that explicitly lists:
  - wrapped now
  - planned next
  - intentionally deferred
  - not in scope unless approved

- [ ] Add one focused example and one focused test per new backlog item

- [ ] Keep `README.MD`, `OVERVIEW.md`, and `TODO.MD` aligned with whichever backlog items actually ship
