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
- `PSWriteOfficeHashtable` for `Import-OfficeCsv -AsHashtable` read lanes
- `PSWriteOfficeDataTable` for `Import-OfficeCsv -AsDataTable` read lanes
- `NativeCsv`

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
    -Engine PSWriteOffice,PSWriteOfficeHashtable,PSWriteOfficeDataTable,NativeCsv `
    -UpdateReadme
```

<!-- BENCHMARK:CsvComparison:START -->
| Scenario | Rows | PSWriteOffice | NativeCsv | PSWriteOfficeDataTable | PSWriteOfficeHashtable | Result |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| csv-read-source-mixed | 1000 | 10.6 ms (1.00x) | 9.2 ms (1.15x faster) | 9.0 ms (1.18x faster) | 8.7 ms (1.23x faster) | PSWriteOfficeHashtable fastest; PSWriteOffice 1.23x slower |
| csv-read-source-mixed | 5000 | 16.8 ms (1.00x) | 22.3 ms (1.33x slower) | 14.2 ms (1.18x faster) | 17.5 ms (1.04x slower) | PSWriteOfficeDataTable fastest; PSWriteOffice 1.18x slower |
| csv-read-source-mixed | 10000 | 63.7 ms (1.00x) | 51.8 ms (1.23x faster) | 60.8 ms (1.05x faster) | 45.2 ms (1.41x faster) | PSWriteOfficeHashtable fastest; PSWriteOffice 1.41x slower |
| csv-read-source-multiline | 1000 | 10.4 ms (1.00x) | 16.5 ms (1.59x slower) | 17.3 ms (1.66x slower) | 10.1 ms (1.02x faster) | PSWriteOfficeHashtable fastest; PSWriteOffice 1.02x slower |
| csv-read-source-multiline | 5000 | 20.4 ms (1.00x) | 38.8 ms (1.90x slower) | 28.5 ms (1.40x slower) | 25.3 ms (1.24x slower) | PSWriteOffice fastest |
| csv-read-source-multiline | 10000 | 109.6 ms (1.00x) | 40.3 ms (2.72x faster) | 34.0 ms (3.22x faster) | 42.4 ms (2.58x faster) | PSWriteOfficeDataTable fastest; PSWriteOffice 3.22x slower |
| csv-read-source-quoted | 1000 | 8.4 ms (1.00x) | 10.8 ms (1.28x slower) | 8.5 ms (1.01x slower) | 9.2 ms (1.10x slower) | PSWriteOffice fastest |
| csv-read-source-quoted | 5000 | 46.4 ms (1.00x) | 29.7 ms (1.57x faster) | 33.1 ms (1.40x faster) | 16.2 ms (2.87x faster) | PSWriteOfficeHashtable fastest; PSWriteOffice 2.87x slower |
| csv-read-source-quoted | 10000 | 55.2 ms (1.00x) | 67.7 ms (1.23x slower) | 31.8 ms (1.74x faster) | 14.6 ms (3.79x faster) | PSWriteOfficeHashtable fastest; PSWriteOffice 3.79x slower |
| csv-read-source-wide | 1000 | 15.2 ms (1.00x) | 13.7 ms (1.10x faster) | 10.5 ms (1.45x faster) | 34.4 ms (2.27x slower) | PSWriteOfficeDataTable fastest; PSWriteOffice 1.45x slower |
| csv-read-source-wide | 5000 | 78.7 ms (1.00x) | 122.7 ms (1.56x slower) | 23.4 ms (3.37x faster) | 43.1 ms (1.83x faster) | PSWriteOfficeDataTable fastest; PSWriteOffice 3.37x slower |
| csv-read-source-wide | 10000 | 142.4 ms (1.00x) | 232.7 ms (1.63x slower) | 54.1 ms (2.63x faster) | 49.2 ms (2.90x faster) | PSWriteOfficeHashtable fastest; PSWriteOffice 2.90x slower |
| csv-write-mixed | 1000 | 11.2 ms (1.00x) | 13.6 ms (1.21x slower) | Skipped | Skipped | PSWriteOffice fastest |
| csv-write-mixed | 5000 | 18.1 ms (1.00x) | 19.5 ms (1.08x slower) | Skipped | Skipped | PSWriteOffice fastest |
| csv-write-mixed | 10000 | 27.8 ms (1.00x) | 43.6 ms (1.56x slower) | Skipped | Skipped | PSWriteOffice fastest |
| csv-write-multiline | 1000 | 13.2 ms (1.00x) | 13.9 ms (1.05x slower) | Skipped | Skipped | PSWriteOffice fastest |
| csv-write-multiline | 5000 | 14.9 ms (1.00x) | 23.8 ms (1.60x slower) | Skipped | Skipped | PSWriteOffice fastest |
| csv-write-multiline | 10000 | 61.8 ms (1.00x) | 58.8 ms (1.05x faster) | Skipped | Skipped | NativeCsv fastest; PSWriteOffice 1.05x slower |
| csv-write-quoted | 1000 | 14.1 ms (1.00x) | 11.3 ms (1.25x faster) | Skipped | Skipped | NativeCsv fastest; PSWriteOffice 1.25x slower |
| csv-write-quoted | 5000 | 20.1 ms (1.00x) | 20.3 ms (1.01x slower) | Skipped | Skipped | PSWriteOffice fastest |
| csv-write-quoted | 10000 | 64.1 ms (1.00x) | 55.7 ms (1.15x faster) | Skipped | Skipped | NativeCsv fastest; PSWriteOffice 1.15x slower |
| csv-write-wide | 1000 | 17.3 ms (1.00x) | 20.3 ms (1.17x slower) | Skipped | Skipped | PSWriteOffice fastest |
| csv-write-wide | 5000 | 70.1 ms (1.00x) | 67.2 ms (1.04x faster) | Skipped | Skipped | NativeCsv fastest; PSWriteOffice 1.04x slower |
| csv-write-wide | 10000 | 79.0 ms (1.00x) | 107.0 ms (1.35x slower) | Skipped | Skipped | PSWriteOffice fastest |
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
