Import-Module ../../PSWriteOffice.psd1 -Force

$path = "$PSScriptRoot/Range.xlsx"
$workbook = New-OfficeExcel
$sheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace
New-OfficeExcelValue -Worksheet $sheet -Row 1 -Column 1 -Value 'Name'
New-OfficeExcelValue -Worksheet $sheet -Row 1 -Column 2 -Value 'Age'
New-OfficeExcelValue -Worksheet $sheet -Row 2 -Column 1 -Value 'John'
New-OfficeExcelValue -Worksheet $sheet -Row 2 -Column 2 -Value 30
New-OfficeExcelValue -Worksheet $sheet -Row 3 -Column 1 -Value 'Jane'
New-OfficeExcelValue -Worksheet $sheet -Row 3 -Column 2 -Value 25
New-OfficeExcelValue -Worksheet $sheet -Row 4 -Column 1 -Value 'Bob'
New-OfficeExcelValue -Worksheet $sheet -Row 4 -Column 2 -Value 40
Save-OfficeExcel -Workbook $workbook -FilePath $path

$rows = Import-OfficeExcel -FilePath $path -StartRow 2 -EndRow 3 -StartColumn 1 -EndColumn 2
$rows | Format-Table
