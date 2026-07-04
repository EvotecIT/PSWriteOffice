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
        $run.Path = $run.OutputPath + '.csv'
        $run.SourcePath = $run.OutputPath + '.source.csv'
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

    comparison Engine -Baseline PSWriteOffice -Metric MedianMs
    artifacts Json, Csv, Markdown
}
