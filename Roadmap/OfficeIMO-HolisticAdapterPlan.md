# PSWriteOffice Holistic OfficeIMO Adapter Plan

PSWriteOffice should be the PowerShell authoring surface for OfficeIMO, not a mirror of older modules or every low-level engine method. Reusable behavior stays in OfficeIMO. Cmdlets translate PowerShell-friendly parameters into OfficeIMO calls, save artifacts, and return useful objects.

## Principles

- Use approved PowerShell verbs for the primary command surface.
- Do not carry PSWritePDF compatibility aliases. PSWritePDF can be archived with migration guidance to PSWriteOffice.
- Prefer beautiful document composition over raw capability exposure.
- Keep Word, PDF, and Markdown mentally aligned where the format allows it.
- Fix missing reusable behavior in OfficeIMO first, then expose it thinly here.
- Validate with generated artifacts, readback/preflight, and package/module import checks.

## PDF Target Surface

PDF is a first-class document family:

- Build: `New-OfficePdf`, `Save-OfficePdf`, `Get-OfficePdf`.
- Compose: themes, headings, paragraphs, rich inline text/link runs, lists, tables, images, panels, horizontal rules, bookmarks, spacers, row/column layout, page breaks, page setup, backgrounds, background images/shapes, page borders, headers, footers, metadata, watermarks.
- Inspect/read: `Get-OfficePdfInfo`, `Get-OfficePdfPreflight`, `Get-OfficePdfText`, `ConvertTo-OfficePdfMarkdown`, `Get-OfficePdfImage`, `Get-OfficePdfAttachment`.
- Operate on pages: `Join-OfficePdf`, `Split-OfficePdf`, `Copy-OfficePdfPage`, `Remove-OfficePdfPage`, `Move-OfficePdfPage`, `Set-OfficePdfPage -Rotation`.

Future PDF slices should add signatures, encryption, redaction, and richer existing-PDF compliance checks only as OfficeIMO.Pdf supports them.

## Composition Model

Word and PDF can share a rich document-building vocabulary: page setup, sections/page breaks, headings, paragraphs, tables, images, panels/callouts, headers, footers, page numbers, watermarks, metadata, and save/export.

Markdown should keep the same authoring mindset with format-appropriate simplification: front matter, headings, paragraphs, lists/task lists, tables, images, links, callouts, details, code blocks, horizontal rules, TOC, and export. It should not pretend to have true page setup, headers/footers, exact positioning, or reliable page breaks in raw Markdown.

Excel should focus on report and workbook workflows: tables, sheets, charts, pivots, sparklines, formulas, template markers, feature inspection, print setup, and PDF export.

PowerPoint should focus on story/deck workflows: sections, semantic slide layouts, visual layout helpers, fit diagnostics, charts/tables/media, notes, themes, and PDF export.

Visio should start from semantic diagram builders and premium-looking output, not raw coordinate wrappers.

## Current Slices

The initial implementation adds the PDF command family above without PSWritePDF aliases.

The second implementation slice adds PDF sidecar export through native document save flows:

- `New-OfficeWord -PdfPath` and `Save-OfficeWord -PdfPath`
- `New-OfficeExcel -PdfPath` and `Save-OfficeExcel -PdfPath`
- `New-OfficeMarkdown -PdfPath` and `Save-OfficeMarkdown -PdfPath`
- `New-OfficePowerPoint -PdfPath` and `Save-OfficePowerPoint -PdfPath`

The third implementation slice adds generated-PDF font options plus existing-PDF and form operations:

- `New-OfficePdf -DefaultFont`, `-DefaultFontSize`, and embedded TrueType `-FontFamily` options.
- `Add-OfficePdfFormField` / `PdfFormField` for generated PDF forms.
- `Get-OfficePdfFormField` for AcroForm inspection.
- `Set-OfficePdfForm -Field ... [-Flatten]` for filling and optional flattening.
- `ConvertTo-OfficePdfFlatForm` for explicit flat-form conversion with an approved verb.
- `Set-OfficePdfMetadata -Path ... -OutputPath ...` for metadata editing on existing PDFs.
- `Add-OfficePdfStamp` / `PdfStamp` for text and image stamps or watermarks on existing PDFs.

The fourth implementation slice adds generated-document attachments, richer PDF readback, and compliance readiness:

- `Add-OfficePdfAttachment` / `PdfAttachment` for generated PDF embedded files.
- `Get-OfficePdfAttachment` for attachment inspection or extraction.
- `Get-OfficePdfImage` for image resource inspection or extraction.
- `ConvertTo-OfficePdfMarkdown` for logical Markdown readback.
- `Set-OfficePdfCompliance` / `PdfCompliance` for generated PDF profile and groundwork settings.
- `Get-OfficePdfCompliance` for generated document readiness reports before saving.

The fifth implementation slice adds visual document polish for generated PDFs:

- `Add-OfficePdfHorizontalRule` / `PdfHorizontalRule` / `PdfHr` for section dividers.
- `Add-OfficePdfBookmark` / `PdfBookmark` for named destinations in generated PDFs.
- `Set-OfficePdfBackground` / `PdfBackground` for document page background color.
- `Set-OfficePdfPageBorder` / `PdfPageBorder` for document page border decoration.

The sixth implementation slice adds layout rhythm for generated PDFs:

- OfficeIMO.Pdf now exposes `PdfDocument.Row(...)` as the reusable engine entry point for percentage-based row/column layout in normal document flow.
- `Add-OfficePdfSpacer` / `PdfSpacer` / `PdfSpace` add vertical rhythm.
- `Add-OfficePdfRow` / `PdfRow` maps PowerShell-friendly column specifications to OfficeIMO.Pdf row composition for headings, paragraphs, panels, lists, tables, horizontal rules, bookmarks, and spacers.

The seventh implementation slice adds generated-document styling:

- `Set-OfficePdfTheme` / `PdfTheme` applies OfficeIMO.Pdf theme presets such as `WordLike`, `TechnicalDocument`, `Compact`, and `Report`.
- `Set-OfficePdfBackgroundImage` / `PdfBackgroundImage` maps directly to OfficeIMO.Pdf page background images.
- `Add-OfficePdfBackgroundShape` / `PdfBackgroundShape` exposes decorative rectangles, ellipses, and anchored bands owned by OfficeIMO.Pdf.
- `Clear-OfficePdfBackgroundShape` clears generated page background shapes.

The eighth implementation slice adds rich inline text:

- `Add-OfficePdfText` / `PdfText` exposes OfficeIMO.Pdf paragraph runs for bold, italic, underline, strike, color, highlight, font, baseline, URI links, and bookmark links.
- `Add-OfficePdfRow` column specifications can use the same `Run`/`Runs` shape for rich inline text inside row/column layouts.

The old PSWritePDF HTML-to-PDF command should not be recreated in PSWriteOffice until that conversion is an OfficeIMO-owned capability. The next PDF slices should add signatures, encryption, redaction, and richer existing-PDF compliance checks only as OfficeIMO.Pdf supports them.
