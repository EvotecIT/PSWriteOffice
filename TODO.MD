# PSWriteOffice reboot — delivery plan

## Vision
- Single PSWriteOffice module (no split packages) built on OfficeIMO.* as the core engine (Word/Excel/Markdown/CSV/PowerPoint).
- Scriptblock‑first DSL (pswritehtml style) but cmdlet names stay verb‑noun and non‑conflicting: e.g., `New-OfficeWord`, `New-OfficeWordSection`; nested aliases like `WordSection`/`WordParagraph` for brevity inside scriptblocks.
- Respect `Get-Verb` guidance; when verbs are awkward, keep the compliant name and add short aliases for DSL ergonomics.
- Focus order: Writer/Reader for Word → Excel → Markdown first; PowerPoint next; converters (HTML/Markdown/PDF) and Visio last/experimental because APIs are unstable.
- OfficeIMO owns feature logic and tests; PSWriteOffice adds a thin DSL/cmdlet layer. Pester only for stabilized cmdlets to confirm wiring.
- Ignore `ImportExcel-master` in OfficeIMO; we are not copying its API.

## Principles (2026)
- OfficeIMO is the reusable C# core; add feature logic there first.
- PSWriteOffice is a thin binary cmdlet layer only (no PowerShell functions).
- Prefer OfficeIMO APIs that are clean and consistent, even if they diverge from ImportExcel naming.

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
  ConvertFrom-OfficeMarkdown -Path report.md -Out report.docx -Theme 'Light'   # later
  ConvertTo-OfficeMarkdown report.docx | Set-Content report.md                 # later
  ConvertTo-OfficeWordPdf report.docx -Out report.pdf                          # later
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
| Excel | `New-OfficeExcel`, `Get-OfficeExcel`, `Save-OfficeExcel`; `Add-OfficeExcelSheet`, `Set-OfficeExcelCell`, `Add-OfficeExcelTable`, `Set-OfficeExcelNamedRange`, `Add-OfficeExcelValidationList`, `Set-OfficeExcelFormula`, `Set-OfficeExcelHeaderFooter`, `Invoke-OfficeExcelAutofit`, `Add-OfficeExcelImage`, `Set-OfficeExcelHyperlink`, `Add-OfficeExcelPivotTable` | `ExcelSheet`, `ExcelCell`, `ExcelTable`, `ExcelNamedRange`, `ExcelValidationList`, `ExcelFormula`, `ExcelHeaderFooter`, `ExcelAutoFit`, `ExcelImage`, `ExcelHyperlink`, `ExcelPivotTable` |
| PowerPoint (exp) | `New-OfficePowerPoint`, `Get-OfficePowerPoint`, `Save-OfficePowerPoint`; `Add-OfficePowerPointSlide`, `Set-OfficePowerPointSlideTitle`, `Add-OfficePowerPointTextBox`, `Add-OfficePowerPointShape`, `Add-OfficePowerPointTable`, `Remove-OfficePowerPointSlide` | `PptSlide`, `PptTitle`, `PptTextBox`, `PptShape`, `PptTable` (planned) |
| Markdown (later) | `Convert-MarkdownToWord`, `Convert-WordToMarkdown`, `Convert-MarkdownToPdf` (preview), `New-OfficeMarkdown` (if needed) | n/a for now |
| Converters (later/preview) | `ConvertTo-OfficeWordPdf`, `ConvertTo-OfficeWordHtml`, `ConvertFrom-OfficeWordHtml`, `ConvertTo-OfficeWordMarkdown`, `ConvertFrom-OfficeMarkdown` | n/a; guard behind `-EnablePreviewConverters` |
| Visio (exp) | `New-OfficeVisio`, `Add-OfficeVisioPage`, `Add-OfficeVisioShape`, `Add-OfficeVisioConnector`, `Save-OfficeVisio` | `VisioPage`, `VisioShape`, `VisioConnector` (experimental) |

## Work packages (ordered, status)
1) **Foundation (thin layer only)** — **In Progress**
   - Keep single module; no PowerShell functions (cmdlets only). **Done**
   - Naming + alias rules remain `Get-Verb` compliant; short aliases only for DSL. **Done**
   - Context objects for nested scriptblocks (document/section/sheet/slide). **Done**

2) **Quality fixes (P0)** — **Done**
   - Fix `-PassThru` binding on Word Section/Paragraph cmdlets. **Done**
   - Wire `New-OfficeExcel -AutoSave` to OfficeIMO. **Done**
   - Align CSV `AllowEmptyLines` docs with parser behavior. **Done**

3) **Excel parity via OfficeIMO (P0/P1)** — **In Progress**
   - Expose existing OfficeIMO Excel features via cmdlets (AutoFilter, Sort, Freeze, ConditionalFormatting, Charts, Images, Hyperlinks). **In Progress** (AutoFilter/Sort/Freeze/ConditionalFormatting/Charts/Images/Hyperlinks cmdlets added; page setup/margins/gridlines/visibility cmdlets added; validation rule cmdlets added)
   - Add missing OfficeIMO Excel features only where needed (comments/notes, protection, sparklines, pivot). **In Progress** (worksheet protection API + cmdlets, comments/notes API + cmdlets, sparklines API + cmdlet, basic pivot tables API + cmdlet + page fields/styles/layout toggles + read cmdlet)
   - Add pivot table read metadata cmdlet for inspection. **Done** (`Get-OfficeExcelPivotTable` + `ExcelPivotTables`)
   - Ensure cmdlets stay thin wrappers and reuse OfficeIMO types. **Pending**

4) **Word polish (P1)** — **In Progress**
   - Surface missing OfficeIMO Word features in cmdlets (bookmarks, content controls, fields, TOC update knobs). **In Progress** (bookmark cmdlet + content control read/write cmdlets + TOC get/set/update cmdlets added)

5) **PowerPoint (P2, experimental)** — **In Progress**
   - Minimal deck/slide/title/textbox/shape/table done; add notes/bullets/images/layout helpers. **In Progress** (notes/bullets/images cmdlets added; layout list + slide by name/type + placeholder queries added; placeholder text + layout bounds/margins/style helpers added)
   - Keep `-Experimental` gate for advanced functionality. **In Progress**

6) **Markdown/CSV (P2)** — **In Progress**
   - Keep OfficeIMO.Markdown/CSV as core; add thin cmdlets for remaining operations if needed. **In Progress** (Markdown DSL cmdlets + example added)

7) **Converters & Visio (P3)** — **Pending**
   - Converters behind preview flags; Visio remains experimental.

8) **Docs & examples (continuous)** — **In Progress**
- Keep examples aligned with new cmdlets and aliases. **In Progress**
- XML help coverage for new cmdlets. **Pending**
- Usage guide for Word/Excel/PowerPoint/Markdown/CSV. **Done**
- Area guides + cmdlet map. **Done**

### Recent updates
- Added a repo-backed OfficeIMO gap backlog grouped by area and scope (`Roadmap/OfficeIMO-GapBacklog.md`).
- Added PowerPoint placeholder example (`ExamplePowerPoint09-Placeholders.ps1`) and README quick start entries for layout placeholder helpers.
- Added Markdown DSL cmdlets, tests, and example (`Example-MarkdownDsl.ps1`).
- Added a usage guide covering Word/Excel/PowerPoint/Markdown/CSV (`Docs/Usage.md`).
- Added area docs and cmdlet map (`Docs/Word.md`, `Docs/Excel.md`, `Docs/PowerPoint.md`, `Docs/Markdown.md`, `Docs/Csv.md`, `Docs/CmdletMap.md`).
- Added advanced Excel example (`Examples/Excel/Example-ExcelAdvanced.ps1`).
- Added PowerPoint DSL context + docs/test updates and a new DSL example (`ExamplePowerPoint10-Dsl.ps1`).
- Enabled PowerPoint getters inside DSL contexts and added inspection example (`ExamplePowerPoint11-InspectionDsl.ps1`).
- Enabled layout placeholder setters inside the DSL and added a layout editing example (`ExamplePowerPoint12-LayoutDsl.ps1`).
- Added layout placeholder getter alias (`PptLayoutPlaceholders`) with docs/test/example updates (`ExamplePowerPoint13-LayoutPlaceholderAliases.ps1`).
- Expanded Markdown/CSV XML help with multi-table Markdown and schema-focused CSV examples.

## Immediate next steps (for this reboot)
- Finalize cmdlet nouns/aliases for Word/Excel to avoid conflicts and ensure `Get-Verb` compliance. **Done**
- Scaffold Word/Excel writer/reader cmdlets with context objects; stub aliases for DSL use. **Done**
- Port 3–5 Word/Excel examples using new names; add minimal Pester smoke (create/open/read). **Done**
- Draft XML help stubs for new cmdlets; ignore platyPS. **In Progress**
