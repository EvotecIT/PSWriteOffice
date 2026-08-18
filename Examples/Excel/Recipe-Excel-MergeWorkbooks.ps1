param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Excel')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$target = Join-Path $OutputDirectory 'Excel-Consolidated.xlsx'
$source = Join-Path $OutputDirectory 'Excel-Regional-Source.xlsx'

ExcelNew -Path $target {
    ExcelSheet 'Summary' { ExcelCell -Address A1 -Value 'Consolidated service report' }
}
ExcelNew -Path $source {
    ExcelSheet 'North' { ExcelCell -Address A1 -Value 'North region'; ExcelCell -Address B1 -Value 42 }
    ExcelSheet 'South' { ExcelCell -Address A1 -Value 'South region'; ExcelCell -Address B1 -Value 37 }
}

$result = Join-OfficeExcelWorkbook -InputPath $target -SourcePath $source -SourceSheet North,South -SheetNamePrefix 'Region '
$summary = Get-OfficeExcelSummary -Path $target -IncludeSheets

[pscustomobject]@{
    Path          = $target
    SheetsCopied  = $result.SheetCount
    WorkbookSheets = ($summary.Sheets.Name -join ', ')
} | Format-List
