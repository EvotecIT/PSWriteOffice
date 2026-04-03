# PSWriteOffice Issue Release Execution

This document turns the issue backlog into a practical release sequence.

## Guardrails

- Keep `PSWriteOffice` and `OfficeIMO` PRs separate.
- Keep each PR scoped to its dedicated worktree branch.
- Do not use `Fixes`, `Closes`, or `Resolves` in commits or PR descriptions.
- Do not close any issue until the released package is publicly available.

## Recommended PR order

1. `OfficeIMO` `codex/officeimo-word-underline-tabs`
   Scope: `#12`
   Why first: it is a real upstream bugfix with focused coverage and no dependency on PSWriteOffice.
2. `OfficeIMO` `codex/officeimo-word-table-content-aggregates`
   Scope: reader ergonomics for `#3`
   Why second: it is optional polish, but if we want the better document-level aggregates in the first release, PSWriteOffice needs to consume this package.
3. `PSWriteOffice` `codex/pswriteoffice-word-issue-surface`
   Scope: `#1`, `#8`, `#18`, `#26`, plus the PSWriteOffice-level `#15` regression.
4. `PSWriteOffice` `codex/pswriteoffice-word-table-cells`
   Scope: `#3`, `#4`, `#14`
   Note: this branch is already strong on paragraphs, images, lists, and nested tables in cells.
5. `PSWriteOffice` `codex/pswriteoffice-word-charts`
   Scope: `#7`, `#19`, plus chart-in-cell guidance for `#3`
   Note: if the table-cell and chart branches are merged independently, keep the final integration smoke test explicit because both touch Word authoring flows.
6. `PSWriteOffice` `codex/pswriteoffice-issue-docs`
   Scope: `#5`, `#20`, `#21`, release notes, issue matrix, and post-release close guidance.

## Merge dependencies

- `codex/officeimo-word-underline-tabs` is independent.
- `codex/officeimo-word-table-content-aggregates` is independent from the underline fix.
- `codex/pswriteoffice-word-issue-surface` is independent from the table-cell and chart branches.
- `codex/pswriteoffice-word-table-cells` is independent from the chart cmdlet branch for pictures, lists, and nested tables.
- `codex/pswriteoffice-word-charts` provides the supported chart path, including chart-in-cell through `-Paragraph`.
- If we want the best reader experience for table-cell images/charts, consume the `OfficeIMO` aggregate-fix package before the PSWriteOffice release cut.

## Release-candidate smoke tests

Run these after merge and before publishing:

1. `PSWriteOffice` Word DSL tests on the release branch.
2. A manual example run for:
   `Example-WordLineBreaks.ps1`
   `Example-WordTableCells.ps1`
   `Example-WordCharts.ps1`
   `Example-WordTableCalculatedColumns.ps1`
3. One end-to-end smoke test for:
   replace text
   transpose table output
   close current/all document tracking
   list/image/nested table in a cell
   chart anchored to a paragraph inside a table cell
4. If the OfficeIMO aggregate fix is included, verify `Get-OfficeWord` read-back sees body table images/charts through document-level collections.

Practical runner:

- `pwsh -File .\Roadmap\Invoke-IssueReleaseSmokeTests.ps1 -FailOnMissingArtifacts`

## Release note checklist

- Mention Word search/replace support.
- Mention transpose support for Word tables.
- Mention expanded Word table layout options.
- Mention improved `Close-OfficeWord` cleanup behavior.
- Mention table-cell authoring support for paragraphs, lists, images, nested tables, and chart anchoring.
- Mention first-class Word chart cmdlet support.
- Mention docs/examples for line breaks, calculated table columns, and chart migration.
- Mention any consumed `OfficeIMO` fixes, especially underline tabs and body table aggregate readers.

## Post-release close order

1. Close `#10` and `#13`.
2. Close docs/example issues: `#5`, `#19`, `#20`, `#21`.
3. Close direct PSWriteOffice wrapper issues: `#1`, `#7`, `#8`, `#18`, `#26`.
4. Close table-cell issues: `#3`, `#4`, `#14`.
5. Close upstream-dependent issues only after confirming the released PSWriteOffice version consumes the right OfficeIMO package:
   `#12`
   `#15`

## If we want to defer one thing

The safest deferrable polish is `codex/officeimo-word-table-content-aggregates`.
The feature paths for authoring already work, and this branch only improves document-level read-back ergonomics for body table images and charts.
