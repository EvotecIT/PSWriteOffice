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

Compare only selected engines:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Standard -Engine PSWriteOffice,ExcelFast
```

## Scenario Suites

`Smoke` is a quick confidence pass for default and table export/import paths.

`Standard` covers the everyday decisions people make: default export, table export, no-table export, autofit, full-sheet import, range import, wide objects, and DataTable input.

`Large` runs the broad workflow family at `25k`, `100k`, and `250k` rows, including the PSWriteOffice DataSet worksheet path.

`Full` includes everything in `Standard` plus a `100k` row count, more repeats, and PSWriteOffice-only DataSet worksheet export.

`SuperLarge` runs scale-safe workflows at `250k`, `500k`, and `1m` rows. It intentionally skips table/autofit/DataSet defaults; use `-Scenario` and `-RowCount` when you want to force those expensive paths.

## Output Files

Every run writes these files:

- `excel-performance-comparison.csv`: one row per scenario/row count with fastest engine, PSWriteOffice rank, PSWriteOffice vs fastest text, competitor timings, and file sizes.
- `excel-performance-comparison.json`: nested comparison data with per-engine rank, timing ratio, and file-size ratio.
- `excel-performance-summary.csv`: median/min/max data grouped by engine and scenario.
- `excel-performance-results.csv`: raw per-iteration results, including failures.
- `metadata.json`: tool versions, selected suite, engines, filters, and output paths.

For quick reading, start with `excel-performance-comparison.csv`.

## Notes

The benchmark records failures as rows in `excel-performance-results.csv` instead of hiding them. That makes unsupported competitor scenarios visible without stopping the whole run.

Install behavior is controlled by `-SkipImportExcelInstall` and `-SkipExcelFastInstall`. Without those switches, missing competitor modules are saved into the benchmark module cache under `Ignore`.
