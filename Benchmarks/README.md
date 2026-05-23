# Excel Benchmarks

`Compare-ExcelPerformance.ps1` compares PSWriteOffice against ImportExcel and ExcelFast across common workbook workflows. It writes raw results, a summary, and metadata under `Ignore\Benchmarks\ExcelPerformance\Run-*`.

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

Longer comparison run:

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Full
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

`Full` includes everything in `Standard` plus larger default row counts and PSWriteOffice-only DataSet worksheet export.

## Notes

The benchmark records failures as rows in `excel-performance-results.csv` instead of hiding them. That makes unsupported competitor scenarios visible without stopping the whole run.

Install behavior is controlled by `-SkipImportExcelInstall` and `-SkipExcelFastInstall`. Without those switches, missing competitor modules are saved into the benchmark module cache under `Ignore`.
