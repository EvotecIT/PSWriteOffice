# PSWriteOffice Issue Release Backlog

This document maps the current open GitHub issues to the work needed in `PSWriteOffice` and `OfficeIMO`.

Related documents:

- `Roadmap/Issue-Release-Execution.md`
- `Roadmap/Issue-Close-Templates.md`

It is intentionally release-safe:

- Do not close any GitHub issue until the first released version containing the fix is published.
- In PRs and commits, use `Refs #<issue>` instead of `Fixes`, `Closes`, or `Resolves`.
- Treat "docs/examples only" issues as closable only after the improved guidance is released.

## Working branches and worktrees

| Repo | Worktree | Branch | Purpose |
| --- | --- | --- | --- |
| `PSWriteOffice` | `C:\Support\GitHub\_wt\pswriteoffice-word-issue-surface` | `codex/pswriteoffice-word-issue-surface` | Functional cmdlet work for Word issue surface |
| `PSWriteOffice` | `C:\Support\GitHub\_wt\pswriteoffice-word-table-cells` | `codex/pswriteoffice-word-table-cells` | Table-cell composition work for Word |
| `PSWriteOffice` | `C:\Support\GitHub\_wt\pswriteoffice-word-charts` | `codex/pswriteoffice-word-charts` | First-class Word chart cmdlet surface |
| `PSWriteOffice` | `C:\Support\GitHub\_wt\pswriteoffice-issue-docs` | `codex/pswriteoffice-issue-docs` | Issue docs, examples, release notes, repo hygiene |
| `OfficeIMO` | `C:\Support\GitHub\_wt\officeimo-word-issue-validation` | `codex/officeimo-word-issue-validation` | Validation tests and upstream fixes only when proven necessary |
| `OfficeIMO` | `C:\Support\GitHub\_wt\officeimo-word-underline-tabs` | `codex/officeimo-word-underline-tabs` | PR-sized upstream fix for underline + tab round-tripping |
| `OfficeIMO` | `C:\Support\GitHub\_wt\officeimo-word-table-content-aggregates` | `codex/officeimo-word-table-content-aggregates` | PR-sized upstream fix for body table image/chart document aggregates |

## Issue triage summary

### Close after release if no further regressions are found

- `#10` Set font name
  Current `OfficeIMO` tests show `FontFamily` also drives HighAnsi/EastAsia/ComplexScript as expected.
- `#13` Header / Footer Support
  `PSWriteOffice` now has first-class header/footer cmdlets and docs.

### Close after release once released docs/examples ship

- `#5` Add Line Breaks
  Modern example now exists for paragraph breaks and same-paragraph breaks.
- `#19` PieChart method
  Updated example exists, and a dedicated Word chart cmdlet now backs it up.
- `#20` Add some extra columns!?
  Example now shows object projection/property-driven table columns.
- `#21` License
  `LICENSE` file and manifest `LicenseUri` cleanup are in place on the docs branch.

### Implemented and verify again at release cut

- `#1` Search and Replace Text
  `Update-OfficeWordText` / `Replace-OfficeWordText` now wrap `OfficeIMO.Word.WordDocument.FindAndReplace`.
- `#3` Add Chart/Picture into Word Table Cell
  Table-cell DSL now supports paragraphs, images, lists, and nested tables.
  Charts can now be anchored to a paragraph created inside a table cell by using `Add-OfficeWordChart -Paragraph $cell.AddParagraph()`.
  Reader note: a dedicated `OfficeIMO` branch now fixes the top-level `WordDocument.Images` / `WordDocument.Charts` aggregates for body table content.
- `#4` Add table to table cell
  Nested tables are now exposed from the PowerShell DSL.
- `#7` No Charts :(
  `Add-OfficeWordChart` / `WordChart` now provide a PowerShell-first chart path.
- `#8` -Transpose parameter
  `Add-OfficeWordTable` now has transpose support.
- `#14` Bulleted list within table cell?
  Lists inside table cells are now covered by the table-cell DSL.
- `#18` Close-OfficeWord unable to close an open worddoc without a valid WordDoc object being passed.
  Document tracking plus current/all cleanup ergonomics are implemented.
- `#26` TableLayout AutoFit behavior, no AutoFitContent or AutoFitWindow?
  Table layout modes now map cleanly to the upstream options.

### Validate first in OfficeIMO, then close only after released package uptake

- `#12` Underline spaces/tabs
  Focused regression test exists and a PR-sized `OfficeIMO` fix is ready on its own worktree.
- `#15` Null array error when exporting to word
  `OfficeIMO` validation and PSWriteOffice-level regression coverage both suggest this is no longer an active corruption bug.
  Keep it open until the released package includes that coverage and we do one final smoke test.

## Planned delivery by branch

### `codex/pswriteoffice-word-issue-surface`

- `#1` Replace cmdlet
- `#8` Transpose support
- `#18` Close-OfficeWord ergonomics
- `#26` Expanded table layout modes

Implemented and verified on this branch:

- `dotnet build Sources\PSWriteOffice.sln` passed
- `Invoke-Pester Tests\WordDsl.Tests.ps1 -Output Detailed` passed with replacement, transpose, close-tracking, and null-row regression coverage

### `codex/pswriteoffice-word-table-cells`

- `#3` Images/charts in table cells
- `#4` Nested tables
- `#14` Lists in table cells
- Supporting DSL/host plumbing as needed

Implemented and verified on this branch:

- `WordTableCell` cmdlet and cell-aware host plumbing
- Paragraphs, lists, images, and nested tables inside cells
- `Invoke-Pester Tests\WordDsl.Tests.ps1 -Output Detailed` passed with cell-composition coverage

Still open on this branch:

- Merge or consume the upstream aggregate fix if we want the document-level readers to include body table images/charts in the released package

### `codex/pswriteoffice-issue-docs`

- `#5` Line-break examples
- `#7` Word chart examples/guidance
- `#19` Pie-chart migration guidance
- `#20` Extra-column table guidance
- `#21` License file and manifest cleanup
- Release notes issue matrix for post-release closure

Implemented on this branch so far:

- `#5` `Examples/Word/Example-WordLineBreaks.ps1`
- `#19` `Examples/Word/Example-WordCharts.ps1`
- `#20` `Examples/Word/Example-WordTableCalculatedColumns.ps1`
- `#21` `LICENSE` and manifest `LicenseUri` cleanup
- Release-safe issue matrix and close-order guidance

### `codex/pswriteoffice-word-charts`

- `#7` First-class `Add-OfficeWordChart` / `WordChart`
- `#19` Chart migration path backed by a real cmdlet
- `#3` Supported chart-in-cell path through `-Paragraph` anchored to a table-cell paragraph
- `Invoke-Pester Tests\WordDsl.Tests.ps1 -Output Detailed` passed with direct, DSL, and table-cell chart coverage

### `codex/officeimo-word-issue-validation`

- `#12` Regression test coverage
- `#15` Repro/validation coverage
- Upstream fixes only if tests prove they are still broken

Implemented on this branch so far:

- `#12` targeted underline/tab regression coverage
- `#15` null-entry table export validation coverage

### `codex/officeimo-word-underline-tabs`

- `#12` PR-sized upstream fix isolated to `WordParagraph` plus the focused test
- `dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj --framework net8.0 --filter "FullyQualifiedName~OfficeIMO.Tests.Word.Test_UnderlinedTextWithTabs_UsesTabCharactersAndPreservesDocument"` passed

### `codex/officeimo-word-table-content-aggregates`

- Optional upstream polish for `#3` reader ergonomics
- Makes `WordDocument.Images`, `WordDocument.Charts`, `ParagraphsImages`, and `ParagraphsCharts` include body table content after reload
- Keeps header/footer aggregate behavior unchanged
- `dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj --framework net8.0 --filter "FullyQualifiedName~OfficeIMO.Tests.Word.Test_CreatingWordDocumentWithImagesInTable|FullyQualifiedName~OfficeIMO.Tests.Word.Test_ChartsInTableCells_AppearInDocumentAggregatesAfterReload"` passed

## Release gate for issue closure

Only close an issue when all of the following are true:

1. The code or documentation change is merged.
2. The containing version is released publicly.
3. The release notes mention the relevant behavior change.
4. We can answer the original issue with a short, concrete example or command path.

## Suggested post-release close order

1. Close the clearly shipped issues first: `#10`, `#13`, docs-only items, and any finished cmdlet wrappers.
2. Close `OfficeIMO`-dependent issues only after the released `PSWriteOffice` build consumes the validated upstream package.
3. Keep any issue open if the fix exists only in source, example code, or unreleased package versions.
