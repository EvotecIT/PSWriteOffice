$Path = Join-Path $PSScriptRoot 'AutoSizeFreeze.xlsx'
New-Item -Path $Path -ItemType File -Force | Out-Null
$Data = 1..3 | ForEach-Object { [PSCustomObject]@{ Name = "Row $_"; Value = "Some very long value $_" } }
$Data | Export-OfficeExcel -FilePath $Path -AutoSize -FreezeTopRow -FreezeFirstColumn
