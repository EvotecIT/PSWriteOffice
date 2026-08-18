param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Excel')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$path = Join-Path $OutputDirectory 'Excel-Read-And-Filter.xlsx'
$rows = @(
    [pscustomobject]@{ Service = 'Identity'; Owner = 'IAM'; Incidents = 1; Status = 'Ready' }
    [pscustomobject]@{ Service = 'Messaging'; Owner = 'Collaboration'; Incidents = 5; Status = 'Review' }
    [pscustomobject]@{ Service = 'Files'; Owner = 'Storage'; Incidents = 0; Status = 'Ready' }
)

ExcelNew -Path $path {
    ExcelSheet 'Services' { ExcelTable -Data $rows -TableName 'ServiceHealth' -AutoFit }
}

$imported = @(Import-OfficeExcel -Path $path -WorksheetName Services)
$needsReview = @($imported | Where-Object { $_.Status -eq 'Review' -or $_.Incidents -ge 3 })

[pscustomobject]@{
    Path        = $path
    Imported    = $imported.Count
    NeedsReview = $needsReview.Count
    Services    = ($needsReview.Service -join ', ')
} | Format-List
