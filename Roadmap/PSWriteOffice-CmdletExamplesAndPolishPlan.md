# PSWriteOffice Cmdlet Examples and Polish Plan

Date: 2026-06-06

This plan turns the current broad PSWriteOffice surface into a clearer, more
teachable module without faking engine capability. OfficeIMO remains the owner
of reusable document behavior. PSWriteOffice should expose stable OfficeIMO APIs
through thin, PowerShell-friendly cmdlets, examples, generated help, and
artifact-producing showcases.

## Current Inventory

Source inventory from `Sources/PSWriteOffice/Cmdlets`:

| Family | Cmdlets |
| --- | ---: |
| CSV | 3 |
| Excel | 79 |
| Markdown | 24 |
| PDF | 46 |
| PowerPoint | 49 |
| Word | 79 |
| Total | 280 |

Generated markdown docs currently exist for every compiled cmdlet. The gap is
example depth, not missing doc files.

XML example audit:

| Family | Cmdlets | No XML examples | No contextual XML example | Has multiline XML example |
| --- | ---: | ---: | ---: | ---: |
| CSV | 3 | 0 | 0 | 1 |
| Excel | 79 | 1 | 29 | 2 |
| Markdown | 24 | 0 | 15 | 1 |
| PDF | 46 | 34 | 37 | 8 |
| PowerPoint | 49 | 9 | 36 | 3 |
| Word | 79 | 11 | 43 | 0 |
| Total | 280 | 55 | 160 | 15 |

The audit treats an example as contextual when it shows a realistic call in a
workflow, not just a one-line command without surrounding setup. The heuristic
is intentionally conservative, so the next work should review examples by human
usefulness instead of only chasing the numbers.

## Example Quality Bar

Every cmdlet should have at least one relevant XML documentation example in the
source file that feeds generated help. Standalone scripts under `Examples/`
remain the right home for larger narratives, but cmdlet help should not depend
on users finding those scripts first.

Use these rules:

- Use one-line examples only for genuinely atomic commands where setup would be
  noise, such as a simple getter over an already-created object.
- Use multiline examples for lifecycle, artifact, composition, migration,
  layout, style, readback, or destructive/update commands.
- Prefer examples that show real input data, a deterministic output path under
  `Examples/Documents`, save/readback where useful, and the surrounding command
  sequence needed to understand the cmdlet.
- Keep examples thin. They should demonstrate calling PSWriteOffice and
  OfficeIMO-owned behavior, not copy reusable logic into sample code.
- Keep aliases out of primary examples unless the example is explicitly showing
  the DSL shorthand. Use approved verbs and canonical cmdlet names first.
- Regenerate docs after source XML edits and spot-check generated markdown help.

Example work should move by family, not as a 280-cmdlet mega-edit:

1. PDF examples first, because the PDF surface is new and has the largest
   no-example gap.
2. PowerPoint examples next, because many commands need workflow context and the
   next implementation work is PowerPoint-heavy.
3. Word examples next, because existing examples are mostly one-line and report
   composition needs a stronger mental model.
4. Excel and Markdown examples last, except for commands touched by feature PRs.

## Missing Product Work

### 1. PDF Table and Style Depth

This is the best next implementation PR.

OfficeIMO.Pdf already owns rich table styling through `PdfTableStyle` and
`TableStyles`, including presets, borders, fills, stripes, separators, cell
padding, row heights, alignment, vertical alignment, auto-fit, column widths,
weighted widths, min/max widths, captions, footer/header rows, numeric
alignment, and Word-style table presets.

PSWriteOffice currently maps `Add-OfficePdfTable` to data/property/header/align
only. The next wrapper should add a PowerShell-friendly style surface without
creating a PDF table DSL:

- `-Style` accepting an existing `PdfTableStyle` or hashtable.
- `-StylePreset` for OfficeIMO presets such as `Light`, `Minimal`,
  `RightAlignedNumbers`, `TechnicalDocument`, `Compact`, and `Report`.
- `-WordTableStyle` for OfficeIMO-supported Word table style names.
- Direct simple parameters for common report use: header fill/text color, row
  stripe fill, border color/width, row separator, cell padding, spacing before
  and after, font size, header font size, line height, max width, caption, and
  numeric right alignment.
- Column parameters for alignments, vertical alignments, fixed widths,
  min/max widths, and width weights.
- Reuse the same builder from `Add-OfficePdfTable` and row-column table content
  so `PdfRow` table blocks do not become a second style path.

Validation should include generated artifact proof plus readback/preflight where
available. Tests should prove parameter-to-OfficeIMO mapping and at least one
real generated PDF containing styled tables.

### 2. PowerPoint Layout Polish

OfficeIMO.PowerPoint already owns much of the layout behavior now listed as a
PSWriteOffice wrapper gap: shape alignment, distribution, fixed spacing,
stacking, resizing, fit-to-bounds, z-order movement, duplication, grouping,
ungrouping, slide hidden state, slide reordering, grid/view settings, chart
formatting, and table-cell formatting.

The next PowerPoint PRs should expose these as thin commands:

- Shape layout: align, distribute, distribute with fixed spacing, stack, grid,
  fit to bounds/slide/content, and resize.
- Shape arrangement: duplicate, bring forward/back, bring to front, send to
  back, group, and ungroup.
- Slide maintenance: move/reorder slides, hide/show slides, and include these
  states in inspection summaries.
- Chart formatting: title, legend, axis, scale, gridlines, number format,
  series fill/line/marker, trendline, and data label wrappers matching the
  existing Excel mental model where practical.
- Table formatting: padding, horizontal/vertical alignment, borders, row
  heights, merged cells, and preset styles only where OfficeIMO exposes stable
  APIs.
- Diagnostics: expose richer fit/layout warnings already produced by deck plan
  and designer flows, then add engine work only when diagnostics are missing in
  OfficeIMO itself.

Do not add PowerShell-side coordinate engines. The cmdlets should select shapes,
bind friendly units, call OfficeIMO methods, and return updated shapes/slides.

### 3. Word Report Polish

Word already has broad primitive coverage, but the report-composition layer is
still uneven. The next Word work should be small helpers around OfficeIMO-owned
behavior:

- Compact paragraph/run style builders where OfficeIMO already supports the
  style state.
- Table row and column mutation helpers if OfficeIMO owns row/column insert,
  remove, size, merge, split, and style operations.
- Text box helpers and richer shape templates only after confirming the engine
  API is stable.

Avoid a giant Word DSL. The target is a cleaner report authoring vocabulary that
matches PDF and Markdown concepts where the formats overlap.

### 4. Migration and Cookbook Docs

No PSWritePDF compatibility aliases should be added.

The public migration story should be a docs/examples PR, ideally after PDF table
style depth so the examples are attractive:

- Add a PSWritePDF-to-PSWriteOffice migration page.
- Include cookbook examples for create PDF, merge, split, stamp, metadata,
  fill/flatten forms, extract text, extract images, and extract attachments.
- Explicitly mark HTML-to-PDF, signatures, encryption, and redaction as
  OfficeIMO.Pdf engine-first backlog unless stable OfficeIMO APIs exist.
- Archive PSWritePDF with a clear "use PSWriteOffice now" path.

### 5. Excel Confidence

Excel is mostly wrapped. The serious remaining work is proof:

- Confirm pivot and sparkline desktop-open compatibility in OfficeIMO first.
- Keep PSWriteOffice wrappers thin once OfficeIMO confidence exists.
- Add stronger flagship/report examples only after compatibility is proven.
- Keep performance work inside normal OfficeIMO save paths.

## Recommended PR Order

1. PDF table/style depth plus PDF XML examples for the touched commands.
2. PDF example sweep for the remaining PDF commands and PSWritePDF migration
   cookbook skeleton.
3. PowerPoint layout/arrangement wrappers with examples and generated artifact
   proof.
4. PowerPoint chart/table/slide-maintenance wrappers with examples.
5. Word report polish: style helpers, table mutation helpers, and contextual XML
   examples for report commands.
6. Excel pivot/sparkline confidence and flagship showcase updates.
7. Full generated-help cleanup pass to reduce the remaining contextual-example
   gaps family by family.

## Validation Expectations

For implementation PRs:

- Build both target frameworks from `Sources/PSWriteOffice.sln`.
- Run focused Pester tests for the touched family.
- Run `Build/Validate-PackagedArtefact.ps1` before calling the module package
  healthy.
- For source-built checks, set `PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES=true`.
- For packaged checks, import the unpacked artifact manifest.
- Generate real example artifacts under `Examples/Documents`.
- Use readback/preflight/inspection summaries when the format supports it.
- Regenerate docs from XML and verify the changed docs contain the intended
  multiline examples.

For docs-only/example PRs:

- Build enough to regenerate help.
- Run the example scripts touched by the PR.
- Verify deterministic outputs are written under `Examples/Documents`.
- Do not add low-value tests that only prove an old command or old module name
  is absent; absence tests are useful only for current public contracts such as
  "PSWritePDF aliases must not be exported."
