$data = 1..5 | ForEach-Object { [PSCustomObject]@{ Value = $_ } }
$chart = @{ Title = 'Chart1'; Range = 'A1:B6' }
$path = Join-Path $PSScriptRoot 'Chart.xlsx'
$data | Export-OfficeExcel -FilePath $path -WorksheetName 'Data' -Charts $chart -Show
