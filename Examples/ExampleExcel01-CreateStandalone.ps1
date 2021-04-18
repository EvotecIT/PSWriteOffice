Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Excel = New-OfficeExcel -FilePath "C:\Support\GitHub\PSWriteOffice\Examples\Documents\Excel2.xlsx"
$Worksheet = Get-OfficeExcelWorkSheet -Excel $Excel -Name 'Contact1'
if (-not $Worksheet) {
    $Worksheet = New-OfficeExcelWorksheet -Excel $Excel -Name 'Contact1'
}
$Cell = Get-OfficeExcelValue -Worksheet $Worksheet -Row 2 -Column 6
if ($Cell) {
    Write-Color "Good" -Color Green
}
Save-OfficeExcel -Excel $Excel -Show