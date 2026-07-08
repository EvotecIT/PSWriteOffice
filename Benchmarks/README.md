# PSWriteOffice Benchmarks

PSWriteOffice benchmarks use the PSPublishModule/PowerForge benchmark DSL. The
suite is split by file format so workbook behavior is compared with workbook
tools and CSV behavior is compared with CSV tools.

## Excel

`Compare-ExcelPerformance.ps1` measures OfficeIMO-backed PSWriteOffice Excel
cmdlets against PowerShell-facing Excel alternatives:

- `PSWriteOffice`
- `ImportExcel`
- `ExcelFast` for the workbook lanes it supports

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Smoke
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Standard -ListScenarios
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 `
    -Suite Standard `
    -RowCount 1000,5000,10000 `
    -RepeatCount 2 `
    -Engine PSWriteOffice,ImportExcel,ExcelFast `
    -UpdateReadme
```

<!-- BENCHMARK:ExcelComparison:START -->
| Scenario | Rows | PSWriteOffice | ExcelFast | ImportExcel | Result |
| --- | ---: | ---: | ---: | ---: | --- |
| append-existing-table | 1000 | 74.8 ms (1.00x) | Skipped | 344.6 ms (4.61x slower) | PSWriteOffice fastest |
| append-existing-table | 5000 | 443.0 ms (1.00x) | Skipped | 1.16 s (2.61x slower) | PSWriteOffice fastest |
| append-existing-table | 10000 | 931.8 ms (1.00x) | Skipped | 2.47 s (2.65x slower) | PSWriteOffice fastest |
| chart-only-workbook | 1000 | 113.8 ms (1.00x) | Skipped | 422.5 ms (3.71x slower) | PSWriteOffice fastest |
| chart-only-workbook | 5000 | 626.1 ms (1.00x) | Skipped | 1.09 s (1.74x slower) | PSWriteOffice fastest |
| chart-only-workbook | 10000 | 1.34 s (1.00x) | Skipped | 2.24 s (1.68x slower) | PSWriteOffice fastest |
| csv-to-excel | 1000 | 78.4 ms (1.00x) | Skipped | 440.0 ms (5.61x slower) | PSWriteOffice fastest |
| csv-to-excel | 5000 | 406.9 ms (1.00x) | Skipped | 1.86 s (4.58x slower) | PSWriteOffice fastest |
| csv-to-excel | 10000 | 737.4 ms (1.00x) | Skipped | 3.52 s (4.78x slower) | PSWriteOffice fastest |
| datatable-default | 1000 | 22.4 ms (1.00x) | Skipped | 254.1 ms (11.36x slower) | PSWriteOffice fastest |
| datatable-default | 5000 | 31.7 ms (1.00x) | Skipped | 819.9 ms (25.88x slower) | PSWriteOffice fastest |
| datatable-default | 10000 | 47.2 ms (1.00x) | Skipped | 1.47 s (31.26x slower) | PSWriteOffice fastest |
| import-default-full | 1000 | 26.4 ms (1.00x) | 45.1 ms (1.71x slower) | 125.2 ms (4.74x slower) | PSWriteOffice fastest |
| import-default-full | 5000 | 117.0 ms (1.00x) | 331.9 ms (2.84x slower) | 321.9 ms (2.75x slower) | PSWriteOffice fastest |
| import-default-full | 10000 | 286.4 ms (1.00x) | 391.2 ms (1.37x slower) | 453.8 ms (1.58x slower) | PSWriteOffice fastest |
| import-default-range | 1000 | 16.3 ms (1.00x) | 34.6 ms (2.13x slower) | 100.4 ms (6.17x slower) | PSWriteOffice fastest |
| import-default-range | 5000 | 75.5 ms (1.00x) | 239.8 ms (3.18x slower) | 237.6 ms (3.15x slower) | PSWriteOffice fastest |
| import-default-range | 10000 | 231.5 ms (1.00x) | 357.8 ms (1.55x slower) | 415.2 ms (1.79x slower) | PSWriteOffice fastest |
| many-small-sheets | 1000 | 120.1 ms (1.00x) | Skipped | 310.3 ms (2.58x slower) | PSWriteOffice fastest |
| many-small-sheets | 5000 | 496.2 ms (1.00x) | Skipped | 1.10 s (2.22x slower) | PSWriteOffice fastest |
| many-small-sheets | 10000 | 976.4 ms (1.00x) | Skipped | 2.13 s (2.18x slower) | PSWriteOffice fastest |
| multi-sheet-regions | 1000 | 150.2 ms (1.00x) | Skipped | 406.6 ms (2.71x slower) | PSWriteOffice fastest |
| multi-sheet-regions | 5000 | 505.7 ms (1.00x) | Skipped | 1.12 s (2.22x slower) | PSWriteOffice fastest |
| multi-sheet-regions | 10000 | 955.5 ms (1.00x) | Skipped | 2.13 s (2.23x slower) | PSWriteOffice fastest |
| named-range-workbook | 1000 | 59.9 ms (1.00x) | Skipped | 271.3 ms (4.53x slower) | PSWriteOffice fastest |
| named-range-workbook | 5000 | 405.0 ms (1.00x) | Skipped | 1.05 s (2.60x slower) | PSWriteOffice fastest |
| named-range-workbook | 10000 | 922.9 ms (1.00x) | Skipped | 2.10 s (2.28x slower) | PSWriteOffice fastest |
| objects-default | 1000 | 80.8 ms (1.00x) | 131.6 ms (1.63x slower) | 300.6 ms (3.72x slower) | PSWriteOffice fastest |
| objects-default | 5000 | 101.7 ms (1.00x) | 177.9 ms (1.75x slower) | 1.01 s (9.94x slower) | PSWriteOffice fastest |
| objects-default | 10000 | 304.6 ms (1.00x) | 367.5 ms (1.21x slower) | 2.52 s (8.27x slower) | PSWriteOffice fastest |
| objects-no-table | 1000 | 31.1 ms (1.00x) | Skipped | 315.8 ms (10.16x slower) | PSWriteOffice fastest |
| objects-no-table | 5000 | 102.5 ms (1.00x) | Skipped | 1.16 s (11.27x slower) | PSWriteOffice fastest |
| objects-no-table | 10000 | 469.5 ms (1.00x) | Skipped | 2.26 s (4.81x slower) | PSWriteOffice fastest |
| objects-table | 1000 | 33.2 ms (1.00x) | Skipped | 310.1 ms (9.33x slower) | PSWriteOffice fastest |
| objects-table | 5000 | 106.6 ms (1.00x) | Skipped | 1.02 s (9.55x slower) | PSWriteOffice fastest |
| objects-table | 10000 | 301.2 ms (1.00x) | Skipped | 2.03 s (6.73x slower) | PSWriteOffice fastest |
| objects-table-autofit | 1000 | 43.1 ms (1.00x) | Skipped | 336.4 ms (7.80x slower) | PSWriteOffice fastest |
| objects-table-autofit | 5000 | 129.1 ms (1.00x) | Skipped | 1.28 s (9.93x slower) | PSWriteOffice fastest |
| objects-table-autofit | 10000 | 341.0 ms (1.00x) | Skipped | 2.30 s (6.75x slower) | PSWriteOffice fastest |
| objects-title-freeze | 1000 | 100.9 ms (1.00x) | Skipped | 360.5 ms (3.57x slower) | PSWriteOffice fastest |
| objects-title-freeze | 5000 | 721.8 ms (1.00x) | Skipped | 1.18 s (1.63x slower) | PSWriteOffice fastest |
| objects-title-freeze | 10000 | 1.05 s (1.00x) | Skipped | 2.07 s (1.97x slower) | PSWriteOffice fastest |
| pivot-only-workbook | 1000 | 115.8 ms (1.00x) | Skipped | 387.2 ms (3.34x slower) | PSWriteOffice fastest |
| pivot-only-workbook | 5000 | 392.5 ms (1.00x) | Skipped | 1.46 s (3.71x slower) | PSWriteOffice fastest |
| pivot-only-workbook | 10000 | 737.3 ms (1.00x) | Skipped | 2.50 s (3.40x slower) | PSWriteOffice fastest |
| read-named-range-metadata | 1000 | 10.8 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-named-range-metadata | 5000 | 60.7 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-named-range-metadata | 10000 | 13.4 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-no-header-range | 1000 | 21.7 ms (1.00x) | 30.7 ms (1.42x slower) | 92.6 ms (4.27x slower) | PSWriteOffice fastest |
| read-no-header-range | 5000 | 87.4 ms (1.00x) | 285.0 ms (3.26x slower) | 251.1 ms (2.87x slower) | PSWriteOffice fastest |
| read-no-header-range | 10000 | 215.5 ms (1.00x) | 311.1 ms (1.44x slower) | 475.3 ms (2.21x slower) | PSWriteOffice fastest |
| read-table-metadata | 1000 | 23.6 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-table-metadata | 5000 | 12.2 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-table-metadata | 10000 | 23.1 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-used-range-datatable | 1000 | 18.2 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-used-range-datatable | 5000 | 97.3 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-used-range-datatable | 10000 | 145.0 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| report-workbook | 1000 | 119.4 ms (1.00x) | Skipped | 358.8 ms (3.00x slower) | PSWriteOffice fastest |
| report-workbook | 5000 | 571.1 ms (1.00x) | Skipped | 1.27 s (2.23x slower) | PSWriteOffice fastest |
| report-workbook | 10000 | 1.48 s (1.00x) | Skipped | 2.10 s (1.41x slower) | PSWriteOffice fastest |
| summary-formulas | 1000 | 62.0 ms (1.00x) | Skipped | 263.0 ms (4.24x slower) | PSWriteOffice fastest |
| summary-formulas | 5000 | 354.7 ms (1.00x) | Skipped | 943.4 ms (2.66x slower) | PSWriteOffice fastest |
| summary-formulas | 10000 | 743.6 ms (1.00x) | Skipped | 1.72 s (2.31x slower) | PSWriteOffice fastest |
| update-existing-workbook | 1000 | 218.3 ms (1.00x) | Skipped | 458.1 ms (2.10x slower) | PSWriteOffice fastest |
| update-existing-workbook | 5000 | 906.7 ms (1.00x) | Skipped | 1.33 s (1.47x slower) | PSWriteOffice fastest |
| update-existing-workbook | 10000 | 1.57 s (1.00x) | Skipped | 2.14 s (1.36x slower) | PSWriteOffice fastest |
| wide-objects-default | 1000 | 77.4 ms (1.00x) | 176.2 ms (2.28x slower) | 242.0 ms (3.13x slower) | PSWriteOffice fastest |
| wide-objects-default | 5000 | 312.6 ms (1.00x) | 345.8 ms (1.11x slower) | 1.03 s (3.29x slower) | PSWriteOffice fastest |
| wide-objects-default | 10000 | 497.9 ms (1.00x) | 680.7 ms (1.37x slower) | 1.66 s (3.33x slower) | PSWriteOffice fastest |
| workbook-package-merge | 1000 | 180.5 ms (1.00x) | Skipped | 608.7 ms (3.37x slower) | PSWriteOffice fastest |
| workbook-package-merge | 5000 | 1.34 s (1.00x) | Skipped | 1.46 s (1.09x slower) | PSWriteOffice fastest |
| workbook-package-merge | 10000 | 1.46 s (1.00x) | Skipped | 2.90 s (1.98x slower) | PSWriteOffice fastest |
<!-- BENCHMARK:ExcelComparison:END -->

## CSV

`Compare-CsvPerformance.ps1` measures PSWriteOffice CSV cmdlets against native
PowerShell CSV import/export:

- `PSWriteOffice`
- `NativeCsv`

The CSV suite also includes dbatools `QuickTest`-shaped read scenarios over the
same 10-column generated CSV shape: `csv-dbatools-quick-single-column` touches
`Column0` for every row, and `csv-dbatools-quick-all-columns` touches every
column for every row.

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-CsvPerformance.ps1 -Suite Smoke
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-CsvPerformance.ps1 -Suite Standard -ListScenarios
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-CsvPerformance.ps1 `
    -Suite Standard `
    -RowCount 1000,5000,10000 `
    -RepeatCount 3 `
    -Engine PSWriteOffice,NativeCsv `
    -UpdateReadme
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-CsvPerformance.ps1 `
    -Suite Smoke `
    -RowCount 100000 `
    -Scenario csv-dbatools-quick-single-column,csv-dbatools-quick-all-columns `
    -Engine PSWriteOffice,NativeCsv
```

Focused local dbatools QuickTest-shaped run, `20260708-053630-00691048`, 100,000
rows, 10 columns, three measured iterations:

| Scenario | PSWriteOffice | NativeCsv | Result |
| --- | ---: | ---: | --- |
| First column read | 341.3 ms, 294,869 rows/s | 347.3 ms, 285,460 rows/s | PSWriteOffice fastest |
| All columns read | 1.08 s, 92,385 rows/s | 1.12 s, 89,942 rows/s | PSWriteOffice fastest |

<!-- BENCHMARK:CsvComparison:START -->
| Scenario | Rows | PSWriteOffice | NativeCsv | Result |
| --- | ---: | ---: | ---: | --- |
| csv-read-source | 1000 | 53.8 ms (1.00x) | 20.8 ms (2.58x faster) | NativeCsv fastest; PSWriteOffice 2.58x slower |
| csv-read-source | 5000 | 61.8 ms (1.00x) | 16.7 ms (3.70x faster) | NativeCsv fastest; PSWriteOffice 3.70x slower |
| csv-read-source | 10000 | 62.2 ms (1.00x) | 179.4 ms (2.89x slower) | PSWriteOffice fastest |
| csv-write | 1000 | 60.0 ms (1.00x) | 51.1 ms (1.17x faster) | NativeCsv fastest; PSWriteOffice 1.17x slower |
| csv-write | 5000 | 124.8 ms (1.00x) | 24.6 ms (5.08x faster) | NativeCsv fastest; PSWriteOffice 5.08x slower |
| csv-write | 10000 | 238.6 ms (1.00x) | 77.3 ms (3.09x faster) | NativeCsv fastest; PSWriteOffice 3.09x slower |
<!-- BENCHMARK:CsvComparison:END -->

## Options

The wrappers build PSWriteOffice in `Release` mode by default and import local
development binaries when a selected run includes `PSWriteOffice`. Use
`-PSWriteOfficeConfiguration Debug` for diagnostics or `-SkipPSWriteOfficeBuild`
when intentionally reusing a previous build. Quick and focused runs leave this
README unchanged unless `-UpdateReadme` is specified.

The scripts use published OfficeIMO packages by default by setting
`OfficeIMORoot` to `.missing-officeimo`. Use `-OfficeIMORoot` when validating
unreleased OfficeIMO source changes:

```powershell
$evotecRoot = if ($env:EVOTEC_GITHUB_ROOT) { $env:EVOTEC_GITHUB_ROOT } else { 'C:\Support\GitHub' }
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 `
    -Suite Standard `
    -OfficeIMORoot (Join-Path $evotecRoot 'OfficeIMO')
```

## Output

Benchmark artifacts are written under `Ignore\Benchmarks\ExcelPerformance` and
`Ignore\Benchmarks\CsvPerformance`:

- `samples.json` / `samples.csv`
- `summary.json` / `summary.csv`
- `comparison.json` / `comparison.md`
- `metadata.json`
- `run-report.json`

Start with `comparison.md` or the generated README table for the comparison,
then use `samples.csv` when diagnosing individual failures.
