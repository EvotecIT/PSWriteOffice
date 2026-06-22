# Excel Benchmarks

`Compare-ExcelPerformance.ps1` compares PSWriteOffice against ImportExcel and ExcelFast across common workbook workflows. It writes raw results, a summary, one-line comparison outputs, and metadata under `Ignore\Benchmarks\ExcelPerformance\Run-*`.

The script uses published OfficeIMO packages by default by setting `OfficeIMORoot` to `.missing-officeimo`, so PSWriteOffice measures the package-mode path instead of a local OfficeIMO checkout.

## Common Runs

List available scenarios:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -ListScenarios
```

Fast sanity run:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Smoke
```

Default comparison run:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Standard
```

Large comparison run:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Large
```

Longer comparison run:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Full
```

Scale stress run:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite SuperLarge
```

Focus on a specific workflow:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Standard -Scenario objects-default -RowCount 25000 -RepeatCount 5
```

Run the richer report workbook workflow:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Standard -Scenario report-workbook -RowCount 1000,10000 -RepeatCount 3 -Engine PSWriteOffice,ImportExcel
```

Run the more operational workbook workflows:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Standard -Scenario objects-title-freeze,multi-sheet-regions,summary-formulas -RowCount 1000,10000,25000 -RepeatCount 3 -Engine PSWriteOffice,ImportExcel
```

Run the append, update, many-sheet, read-focused, and chart/pivot split workflows:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Standard -Scenario objects-default,append-existing-table,update-existing-workbook,many-small-sheets,named-range-workbook,chart-only-workbook,pivot-only-workbook -RowCount 1000,10000,25000 -RepeatCount 3
```

Measure export creation without import follow-up timing:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Standard -Scenario objects-default,wide-objects-default -RowCount 25000 -RepeatCount 3 -Engine PSWriteOffice,ImportExcel,ExcelFast -SkipFollowUps
```

Compare only selected engines:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Standard -Engine PSWriteOffice,ExcelFast
```

## Scenario Suites

`Smoke` is a quick confidence pass for default, table, and report workbook export/import paths.

`Standard` covers the everyday decisions people make: default export, table export, no-table export, autofit, full-sheet import, range import, no-header reads, used-range DataTable reads, table/named-range metadata reads, append-to-existing-table workflows, update-existing-workbook workflows, wide objects, DataTable input, title/start-row/frozen-header exports, regional multi-sheet workbooks, many-small-sheet workbooks, formula summary sheets, chart-only and pivot-only workbooks, and a full report workbook with a table, freeze row, conditional formatting, validation, a chart, and a pivot table.

`Large` runs the broad workflow family at `25k`, `100k`, and `250k` rows, including the PSWriteOffice DataSet worksheet path.

`Full` includes everything in `Standard` plus a `100k` row count, more repeats, and PSWriteOffice-only DataSet worksheet export.

`SuperLarge` runs scale-safe workflows at `250k`, `500k`, and `1m` rows. It intentionally skips table/autofit/DataSet defaults; use `-Scenario` and `-RowCount` when you want to force those expensive paths.

## Output Files

Every run writes these files:

- `excel-performance-comparison.csv`: one row per scenario/row count with fastest engine, per-engine status (`tested`, `failed`, `not selected`, or `not supported by scenario`), PSWriteOffice rank, PSWriteOffice vs fastest text, competitor timings, file sizes, and compact memory deltas.
- `excel-performance-comparison.json`: nested comparison data with per-engine rank, timing ratio, file-size ratio, and memory fields.
- `excel-performance-summary.csv`: median/min/max data grouped by engine and scenario, including median working-set, peak working-set, and managed-memory deltas.
- `excel-performance-results.csv`: raw per-iteration results, including failures, file size, working-set before/after, peak working set, and managed-memory delta.
- `metadata.json`: exact module versions including prerelease labels, machine/runtime details, selected suite, engines, filters, module cache paths, and output paths.

For quick reading, start with `excel-performance-comparison.csv`. See
[Artifact Schema](#artifact-schema) when you need exact column meanings.

## Artifact Schema

`excel-performance-results.csv` is the raw evidence. Important columns:

- `Status` / `Error`: whether the timed operation itself passed.
- `WorkbookValidationStatus`: post-operation workbook validation result for export scenarios.
- `WorkbookOpenStatus`: whether PSWriteOffice could open the generated workbook.
- `WorkbookOpenXmlStatus`: whether Open XML validation passed, failed, or was skipped because validator types were unavailable.
- `WorkbookValidationMs`: validation time, recorded separately from the timed operation.
- `WorkingSetBeforeMB`, `WorkingSetAfterMB`, `WorkingSetDeltaMB`, `PeakWorkingSetDeltaMB`, `ManagedDeltaMB`: memory diagnostics for the operation.

`excel-performance-summary.csv` groups raw rows by engine, scenario, profile,
and row count. It reports median/min/max time, median file size, median memory
deltas, and workbook-validation pass/fail/skip counts.

`excel-performance-comparison.csv` is the comparison view. It includes:

- `FastestEngine`, `FastestMs`, and PSWriteOffice rank/ratio text.
- Per-engine status columns: `tested`, `failed`, `not selected`, or `not supported by scenario`.
- Per-engine workbook validation status and validation time.
- Per-engine median file size and memory deltas.

`metadata.json` records exact module versions including prerelease labels,
machine/runtime details, selected engines/scenarios, module cache paths,
OfficeIMO root, and output paths.

## Notes

The benchmark records failures as rows in `excel-performance-results.csv` instead of hiding them. That makes unsupported competitor scenarios visible without stopping the whole run.

Install behavior is controlled by `-SkipImportExcelInstall` and `-SkipExcelFastInstall`. Without those switches, missing competitor modules are saved into the benchmark module cache under `Ignore`.

Workbook validation is enabled by default for export scenarios. Use
`-SkipWorkbookValidation` only when you need raw timing without post-export
open/OpenXML checks.
