. (Join-Path $PSScriptRoot 'excel-performance.helpers.ps1')

$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$suiteName = input Suite Standard
$rowCounts = Assert-ExcelBenchmarkRowCount -RowCount (inputInt RowCount (Get-ExcelBenchmarkDefaultRowCount -Suite $suiteName))
$skipWorkbookValidation = inputBool SkipWorkbookValidation false
$skipImportExcelInstall = inputBool SkipImportExcelInstall false
$skipExcelFastInstall = inputBool SkipExcelFastInstall false

benchmark 'excel-performance' -out (Join-Path $repositoryRoot 'Ignore\Benchmarks\ExcelPerformance') {
    policy -Warmup 1 -Iterations (Get-ExcelBenchmarkIterationCount -Suite $suiteName) -Order Rotated -OutlierMode None
    profile Current -Cleanup KeepOnFailure
    caseSource (Get-ExcelBenchmarkCase -Suite $suiteName)
    axis RowCount $rowCounts

    setup {
        param($case, $run)

        $run.RepositoryRoot = $repositoryRoot
        $run.WorksheetName = 'Data'
        $run.Path = $run.OutputPath + (Get-ExcelBenchmarkExtension -Case $case)
        $run.SourcePath = $run.OutputPath + '.source.csv'
        $run.SkipWorkbookValidation = $skipWorkbookValidation
        $run.SkipImportExcelInstall = $skipImportExcelInstall
        $run.SkipExcelFastInstall = $skipExcelFastInstall
        $run.Range = Get-ExcelBenchmarkRange -ColumnCount (Get-ExcelBenchmarkColumnCount -Profile $case.DataProfile) -Rows ([int]$case.RowCount)
        $run.RangeEndCell = Get-ExcelBenchmarkRangeEndCell -ColumnCount (Get-ExcelBenchmarkColumnCount -Profile $case.DataProfile) -Rows ([int]$case.RowCount)
        Initialize-ExcelBenchmarkEngine -Engine $case.Engine -Run $run
    }

    Set-BenchmarkDataFactory {
        param($case, $run)

        $profile = Get-ExcelBenchmarkData -Profile $case.DataProfile -Count ([int]$case.RowCount)
        $run.Payload = $profile.Data
        $run.ExpectedRows = [int]$case.RowCount
        $run.ColumnCount = $profile.ColumnCount
        $run.WorksheetName = $profile.WorksheetName
        $run.Range = Get-ExcelBenchmarkRange -ColumnCount $profile.ColumnCount -Rows ([int]$case.RowCount)
        $run.RangeEndCell = Get-ExcelBenchmarkRangeEndCell -ColumnCount $profile.ColumnCount -Rows ([int]$case.RowCount)
        Initialize-ExcelBenchmarkInput -Case $case -Run $run
    }

    skip {
        param($case)

        return ([string] $case.SupportedEngines -split ',') -notcontains [string] $case.Engine
    }

    engine PSWriteOffice {
        operation Run {
            param($case, $run)
            Invoke-ExcelBenchmarkOperation -Engine PSWriteOffice -Case $case -Run $run
        }
    }

    engine ImportExcel {
        operation Run {
            param($case, $run)
            Invoke-ExcelBenchmarkOperation -Engine ImportExcel -Case $case -Run $run
        }
    }

    engine ExcelFast {
        operation Run {
            param($case, $run)
            Invoke-ExcelBenchmarkOperation -Engine ExcelFast -Case $case -Run $run
        }
    }

    validate {
        param($case, $run)

        Test-ExcelBenchmarkOutput -Case $case -Run $run
    }

    comparison Engine -Baseline PSWriteOffice -Metric MedianMs -TieTolerance 0.05 -RequireBaselineFastest
    artifacts Json, Csv, Markdown
}
