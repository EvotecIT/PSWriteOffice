Import-Module ../../PSWriteOffice.psd1 -Force

class Person {
    [string]$Name
    [int]$Age
}

$path = "$PSScriptRoot/Typed.xlsx"
$workbook = New-OfficeExcel
$sheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace
New-OfficeExcelValue -Worksheet $sheet -Row 1 -Column 1 -Value 'Name'
New-OfficeExcelValue -Worksheet $sheet -Row 1 -Column 2 -Value 'Age'
New-OfficeExcelValue -Worksheet $sheet -Row 2 -Column 1 -Value 'John'
New-OfficeExcelValue -Worksheet $sheet -Row 2 -Column 2 -Value 30
Save-OfficeExcel -Workbook $workbook -FilePath $path

$rows = Import-OfficeExcel -FilePath $path -Type ([Person])
$rows | Format-Table
