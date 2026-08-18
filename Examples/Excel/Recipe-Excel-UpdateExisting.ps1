param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Excel')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$path = Join-Path $OutputDirectory 'Excel-Updated-Existing.xlsx'
$rows = @(
    [pscustomobject]@{ Service = 'Identity'; Status = 'Draft'; Owner = 'IAM' }
    [pscustomobject]@{ Service = 'Messaging'; Status = 'Draft'; Owner = 'Collaboration' }
)

ExcelNew -Path $path {
    ExcelSheet 'Readiness' { ExcelTable -Data $rows -TableName 'Readiness' -AutoFit }
}

$replacements = Update-OfficeExcelText -Path $path -Sheet Readiness -OldValue 'Draft' -NewValue 'Ready'
Edit-OfficeExcelRow -Path $path -Sheet Readiness -ScriptBlock {
    param($row)
    if ($row.CellByHeader('Service').Value -eq 'Messaging') {
        $row.Set('Owner', 'Productivity')
    }
}

$updated = @(Import-OfficeExcel -Path $path -WorksheetName Readiness)
[pscustomobject]@{
    Path         = $path
    Replacements = $replacements
    ReadyRows    = @($updated | Where-Object Status -eq Ready).Count
    NewOwner     = ($updated | Where-Object Service -eq Messaging).Owner
} | Format-List
