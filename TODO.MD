# PSWriteOffice reboot — delivery plan

## Vision
- Single PSWriteOffice module (no split packages) that replaces PSWriteWord, PSWriteExcel, ImportExcel, and legacy PSWriteOffice using OfficeIMO as the core engine.
- Scriptblock‑first DSL (pswritehtml style) but cmdlet names stay verb‑noun and non‑conflicting: e.g., `New-OfficeWord`, `New-OfficeWordSection`; nested aliases like `WordSection`/`WordParagraph` for brevity inside scriptblocks.
- Respect `Get-Verb` guidance; when verbs are awkward, keep the compliant name and add short aliases for DSL ergonomics.
- Focus order: Writer/Reader for Word → Excel → Markdown first; PowerPoint next; converters (HTML/Markdown/PDF) and Visio last/experimental because APIs are unstable.
- OfficeIMO owns feature logic and tests; PSWriteOffice adds a thin DSL/cmdlet layer. Pester only for stabilized cmdlets to confirm wiring.

## Guiding DSL shape (prototypes, not final)
- Word
  ```powershell
  New-OfficeWord -Path out.docx {
      WordSection {
          WordHeader { WordParagraph "Quarterly Report" -Bold }
          WordFooter { WordPageNumber }
          WordParagraph { WordText "Hello "; WordBold "world" }
          WordList -Type Bullet {
              WordListItem "One"
              WordListItem "Two"
          }
          WordTable -Data $orders -Style "TableGrid" {
              WordTableCondition -If { $_.Total -gt 1000 } -Style "Accent1"
          }
          WordImage "$PSScriptRoot/logo.png" -Width 96 -Height 32
      }
  } -Open
  ```
- Excel
  ```powershell
  New-OfficeExcel -Path out.xlsx {
      ExcelSheet "Data" {
          ExcelCell A1 "Name" -Bold
          ExcelTable -Data $data -Start A2 -AutoFit -Style "TableStyleMedium9"
          ExcelNamedRange "Totals" "B2:B50"
          ExcelValidationList "C2:C50" -Values "New","In Progress","Done"
      }
      # Charts deferred until OfficeIMO.Excel supports them
  } -Save -Open
  ```
- PowerPoint (behind `-Experimental`)
  ```powershell
  New-OfficePowerPoint -Path out.pptx {
      PptSlide "Title" {
          PptTitle "Status Update"
          PptBullets "Wins","Risks","Next Steps"
          PptImage "$PSScriptRoot/logo.png" -Pos 9,0.5 -Size 2,0.8
          PptNotes "Keep this to 5 minutes."
      }
  } -Experimental
  ```
- Markdown / HTML / PDF
  ```powershell
  Convert-MarkdownToWord -Path report.md -Out report.docx -Theme 'Light'   # later
  Convert-WordToMarkdown report.docx | Set-Content report.md               # later
  Convert-WordToPdf report.docx -Out report.pdf                            # later
  ```
- Reading (object‑first)
  ```powershell
  Get-OfficeExcel -Path data.xlsx -Sheet Data -UsedRange -AsObjects
  Get-OfficeWord -Path report.docx -Sections | Where-Object Header -match "Q4"
  ```

## Cmdlet naming & aliases (plan)
- Canonical cmdlets stay `Verb-Office*` to avoid conflicts; aliases provide the DSL feel inside scriptblocks.
- All verbs follow `Get-Verb` guidance; shortcuts are aliases only.

| Area | Canonical cmdlets (examples) | DSL aliases (used inside scriptblocks) |
| ---- | ---------------------------- | -------------------------------------- |
| Word | `New-OfficeWord`, `Get-OfficeWord`, `Save-OfficeWord`; `Add-OfficeWordSection`, `Add-OfficeWordHeader`, `Add-OfficeWordFooter`, `Add-OfficeWordParagraph`, `Add-OfficeWordText`, `Add-OfficeWordList`, `Add-OfficeWordListItem`, `Add-OfficeWordTable`, `Add-OfficeWordImage`, `Add-OfficeWordField`, `Protect-OfficeWordDocument`, `Find-OfficeWord` | `WordSection`, `WordHeader`, `WordFooter`, `WordParagraph`, `WordText`, `WordBold` (alias to `Add-OfficeWordText -Bold`), `WordList`, `WordListItem`, `WordTable`, `WordTableCondition`, `WordImage`, `WordField`, `WordPageNumber`, `WordWatermark` |
| Excel | `New-OfficeExcel`, `Get-OfficeExcel`, `Save-OfficeExcel`; `Add-OfficeExcelSheet`, `Set-OfficeExcelCell`, `Add-OfficeExcelTable`, `Set-OfficeExcelNamedRange`, `Add-OfficeExcelValidationList`, `Set-OfficeExcelFormula`, `Set-OfficeExcelHeaderFooter`, `Invoke-OfficeExcelAutofit` | `ExcelSheet`, `ExcelCell`, `ExcelTable`, `ExcelNamedRange`, `ExcelValidationList`, `ExcelFormula`, `ExcelHeaderFooter`, `ExcelAutoFit` |
| PowerPoint (exp) | `New-OfficePowerPoint`, `Get-OfficePowerPoint`, `Save-OfficePowerPoint`; `Add-OfficePowerPointSlide`, `Set-OfficePowerPointTitle`, `Add-OfficePowerPointTextBox`, `Add-OfficePowerPointImage`, `Set-OfficePowerPointNotes`, `Add-OfficePowerPointBullet` | `PptSlide`, `PptTitle`, `PptTextBox`, `PptImage`, `PptNotes`, `PptBullets`, `PptBullet` |
| Markdown (later) | `Convert-MarkdownToWord`, `Convert-WordToMarkdown`, `Convert-MarkdownToPdf` (preview), `New-OfficeMarkdown` (if needed) | n/a for now |
| Converters (later/preview) | `Convert-WordToPdf`, `Convert-WordToHtml`, `Convert-HtmlToWord`, `Convert-WordToMarkdown`, `Convert-MarkdownToWord` | n/a; guard behind `-EnablePreviewConverters` |
| Visio (exp) | `New-OfficeVisio`, `Add-OfficeVisioPage`, `Add-OfficeVisioShape`, `Add-OfficeVisioConnector`, `Save-OfficeVisio` | `VisioPage`, `VisioShape`, `VisioConnector` (experimental) |

## Work packages (ordered)
1) **Foundation (module stays single)**
   - Keep PSPublishModule pipeline; only adjust loader to prefer local OfficeIMO sources in dev and packaged binaries for release.
   - Establish naming + alias rules to avoid conflicts and stay `Get-Verb` compliant; ship aliases for DSL brevity.
   - Context objects for nested scriptblocks (document/section/sheet/slide) without leaking .NET classes.

2) **Word (P0)**
   - Writers: New/Open/Save Word, Section/Header/Footer/Paragraph/Text/List/Table/Image/Field/Watermark/Protection; fluent helper aliases inside DSL.
   - Readers: Get-OfficeWord (sections/paragraphs/tables/runs), Find-Word (text/fields/bookmarks).
   - Basic Pester smoke: create doc, add paragraph, add table, save/open roundtrip.
   - Back-compat shims for top PSWriteWord verbs (with deprecation warnings).

3) **Excel (P0)**
   - Writers: New/Open/Save Excel, Sheet, Row/Column/Cell, Table, NamedRange, Validation, Formula, Images in header/footer, AutoFit (hook OfficeIMO parallel). Charts postponed until OfficeIMO.Excel exposes them.
   - Readers: Get-OfficeExcel (-UsedRange/-Range/-AsObjects/-AsRows), Get-ExcelNamedRange/Table helpers; ImportExcel-compatible wrapper.
   - Pester smoke: create workbook, write table from objects, autofit, read back.

4) **Markdown (P1)**
   - Writers/Readers: New-OfficeMarkdown (if needed) or Convert-like commands once stable; start with reading/writing via OfficeIMO.Markdown where stable.
   - Defer HTML/PDF/Markdown converters until OfficeIMO converters stabilize.

5) **PowerPoint (P2, experimental flag)**
   - Minimal deck/slide/title/textbox/bullets/image/notes; gated by `-Experimental`.
   - Reader: Get-OfficePowerPoint to list slides/content (basic).

6) **Converters (P3, later)**
   - Convert-WordTo/From Markdown/Html/Pdf; Convert-MarkdownToWord; wrapped behind stability checks; separate `-EnablePreviewConverters` switch.

7) **Visio (P3, experimental)**
   - Thin DSL for diagrams/shapes/connectors with explicit preview flag.

8) **Examples & docs**
   - Rebuild Examples with new naming (Word/Excel first). Start with “create doc”, “get doc contents”, “create excel table”, “read excel as objects”.
   - XML-based help for cmdlets; platyPS not needed. Include alias coverage in help.
   - Migration guide (old cmdlet → new) once Word/Excel stabilized.

9) **Tests & validation**
   - Trust OfficeIMO unit/verify tests for logic; add Pester smoke/integration only for cmdlet wiring and DSL flow.
   - Cross-platform runnable examples double as tests.

10) **Release & versioning**
    - SemVer: 0.x while DSL incubates; 1.0 when Word/Excel DSL + shims stable.
    - Single module distribution; optional switches to avoid loading heavy deps (PDF/HTML/Markdown) until requested.

## Immediate next steps (for this reboot)
- Finalize cmdlet nouns/aliases for Word/Excel to avoid conflicts and ensure `Get-Verb` compliance.
- Scaffold Word/Excel writer/reader cmdlets with context objects; stub aliases for DSL use.
- Port 3–5 Word/Excel examples using new names; add minimal Pester smoke (create/open/read).
- Draft XML help stubs for new cmdlets; ignore platyPS.
