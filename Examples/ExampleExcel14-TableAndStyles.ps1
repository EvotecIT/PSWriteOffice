$workbook = New-OfficeExcel
$worksheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Sheet1' -Option Replace

$data = @(
    [pscustomobject]@{ Name = 'Item1'; Value = 10 },
    [pscustomobject]@{ Name = 'Item2'; Value = 20 }
)

New-OfficeExcelTable -Worksheet $worksheet -DataTable $data -StartRow 1 -StartColumn 1 -ShowRowStripes -Theme Light9
Set-OfficeExcelCellStyle -Worksheet $worksheet -Row 1 -Column 1 -Bold $true -FontColor 'Red'
Set-OfficeExcelWorkSheetStyle -Excel $workbook -Worksheet $worksheet -TabColor 'Green'

Save-OfficeExcel -Excel $workbook -FilePath "$PSScriptRoot/ExampleExcel14.xlsx" -Show
