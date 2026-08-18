param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Excel')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$baseline = Join-Path $OutputDirectory 'Excel-Baseline.xlsx'
$candidate = Join-Path $OutputDirectory 'Excel-Candidate.xlsx'

ExcelNew -Path $baseline { ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Status'; ExcelCell -Address A2 -Value 'Draft' } }
ExcelNew -Path $candidate { ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Status'; ExcelCell -Address A2 -Value 'Ready' } }

$comparison = Compare-OfficeExcelWorkbook -InputPath $baseline -DifferencePath $candidate
$differences = @($comparison.Differences)

[pscustomobject]@{
    Baseline        = $baseline
    Candidate       = $candidate
    IsEqual         = $comparison.IsEqual
    DifferenceCount = $differences.Count
} | Format-List
