param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Excel')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$csv = Join-Path $OutputDirectory 'Regional-Sales.csv'
$workbook = Join-Path $OutputDirectory 'Excel-Imported-Delimited.xlsx'
Set-Content -Path $csv -Value "Region;Revenue`r`nEMEA;98000.50`r`nAPAC;143000.25" -NoNewline
ExcelNew -Path $workbook { ExcelSheet 'Readme' { ExcelCell -Address A1 -Value 'Imported from a semicolon-delimited source.' } }

$import = Import-OfficeExcelDelimitedText -InputPath $workbook -SourcePath $csv -Delimiter ';' -SheetName Sales -PassThru
$rows = @(Import-OfficeExcel -Path $workbook -WorksheetName Sales)

[pscustomobject]@{
    Path       = $workbook
    Sheet      = $import.SheetName
    Rows       = $rows.Count
    Total      = ($rows | Measure-Object Revenue -Sum).Sum
} | Format-List
