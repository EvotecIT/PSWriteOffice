$data = @(
    [PSCustomObject]@{ A = 1; B = 2 }
    [PSCustomObject]@{ A = 3; B = 4 }
)
$formulas = @{ 'C2' = '=A2+B2'; 'C3' = '=A3+B3' }
$path = Join-Path $PSScriptRoot 'Formulas.xlsx'
$data | Export-OfficeExcel -FilePath $path -WorksheetName 'Data' -Formulas $formulas -Show
