Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Excel = New-OfficeExcel -FilePath "$PSScriptRoot\Documents\Excel3.xlsx" -WhenExists Overwrite
$Worksheet = Get-OfficeExcelWorkSheet -Excel $Excel -Name 'Contact1'
if (-not $Worksheet) {
    $Worksheet = New-OfficeExcelWorksheet -Excel $Excel -Name 'Contact1'
}
$Cell = Get-OfficeExcelValue -Worksheet $Worksheet -Row 2 -Column 6
if ($Cell) {
    Write-Host "Good" -ForegroundColor Green
}

# A tab color to WorkSheet
Set-OfficeExcelWorkSheetStyle -TabColor Red -WorkSheetName 'Contact1' -Excel $Excel

New-OfficeExcelValue -Row 1 -Column 1 -Value 'Test1' -Worksheet $Worksheet
New-OfficeExcelValue -Row 1 -Column 2 -Value 'Test2' -Worksheet $Worksheet
New-OfficeExcelValue -Row 1 -Column 3 -Value 'Test3' -Worksheet $Worksheet
New-OfficeExcelValue -Row 1 -Column 4 -Value 'Test4' -Worksheet $Worksheet
New-OfficeExcelValue -Row 2 -Column 1 -Value 'Test' -Worksheet $Worksheet
New-OfficeExcelValue -Row 2 -Column 2 -Value 'Test' -Worksheet $Worksheet
New-OfficeExcelValue -Row 2 -Column 3 -Value 'Test' -Worksheet $Worksheet
New-OfficeExcelValue -Row 2 -Column 4 -Value 'Test' -Worksheet $Worksheet
$FirstCell = $Worksheet.FirstCell()
$LastCell = $Worksheet.Row(2).Cell(4)
$Range = $Worksheet.Range($FirstCell.Address, $LastCell.Address)
$Range.CreateTable()

Save-OfficeExcel -Excel $Excel -Show
