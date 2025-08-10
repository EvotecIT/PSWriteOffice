$data = @(
    [PSCustomObject]@{ Category = 'A'; Value = 1 }
    [PSCustomObject]@{ Category = 'A'; Value = 2 }
    [PSCustomObject]@{ Category = 'B'; Value = 3 }
)
$pivot = @{ Name = 'Pivot1'; SourceRange = 'A1:B4'; TargetCell = 'D2'; RowFields = @('Category'); Values = @{ Value = 'Sum' } }
$path = Join-Path $PSScriptRoot 'Pivot.xlsx'
$data | Export-OfficeExcel -FilePath $path -WorksheetName 'Data' -PivotTables $pivot -Show
