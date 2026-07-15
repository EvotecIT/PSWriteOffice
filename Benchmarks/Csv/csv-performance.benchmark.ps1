. (Join-Path $PSScriptRoot '..\Excel\excel-performance.helpers.ps1')

$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$suiteName = input Suite Standard
$rowCounts = Assert-ExcelBenchmarkRowCount -RowCount (inputInt RowCount (Get-ExcelBenchmarkDefaultRowCount -Suite $suiteName))

benchmark 'csv-performance' -out (Join-Path $repositoryRoot 'Ignore\Benchmarks\CsvPerformance') {
    policy -Warmup 1 -Iterations (Get-ExcelBenchmarkIterationCount -Suite $suiteName) -Order Rotated -OutlierMode None
    profile Current -Cleanup KeepOnFailure
    caseSource (Get-CsvBenchmarkCase -Suite $suiteName)
    axis RowCount $rowCounts

    setup {
        param($case, $run)

        $run.RepositoryRoot = $repositoryRoot
        $run.WorksheetName = 'Data'
        $extension = Get-ExcelBenchmarkExtension -Case $case
        $run.Path = $run.OutputPath + $extension
        $run.SourcePath = $run.OutputPath + '.source' + $extension
        Initialize-ExcelBenchmarkEngine -Engine $case.Engine -Run $run
    }

    Set-BenchmarkDataFactory {
        param($case, $run)

        $profile = Get-ExcelBenchmarkData -Profile $case.DataProfile -Count ([int]$case.RowCount)
        $run.Payload = $profile.Data
        $run.ExpectedRows = [int]$case.RowCount
        $run.ColumnCount = $profile.ColumnCount
        Initialize-ExcelBenchmarkInput -Case $case -Run $run
    }

    skip {
        param($case)

        -not (Test-CsvBenchmarkEngineSupport -Engine $case.Engine -Case $case)
    }

    engine PSWriteOffice {
        operation Run {
            param($case, $run)
            Invoke-ExcelBenchmarkOperation -Engine PSWriteOffice -Case $case -Run $run
        }
    }

    engine NativeCsv {
        operation Run {
            param($case, $run)
            Invoke-ExcelBenchmarkOperation -Engine NativeCsv -Case $case -Run $run
        }
    }

    validate {
        param($case, $run)

        Test-CsvBenchmarkOutput -Case $case -Run $run
    }

    metric RowsProcessed {
        param($case, $run)

        $run.RowsProcessed
    }

    metric RowsPerSecond {
        param($case, $run)

        if ($run.DurationMs -le 0) {
            return 0
        }

        [double] $case.RowCount / ($run.DurationMs / 1000)
    }

    comparison Engine -Baseline PSWriteOffice -Metric MedianMs -TieTolerance 0.05 -RequireBaselineFastest
    artifacts Json, Csv, Markdown
}
