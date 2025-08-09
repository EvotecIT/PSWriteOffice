Import-Module ../../PSWriteOffice.psd1 -Force

$pathHeader = "$PSScriptRoot/Header.xlsx"
$workbook = New-OfficeExcel
$sheet = New-OfficeExcelWorkSheet -Workbook $workbook -WorksheetName 'Data' -Option Replace
New-OfficeExcelValue -Worksheet $sheet -Row 1 -Column 1 -Value 'Skip'
New-OfficeExcelValue -Worksheet $sheet -Row 1 -Column 2 -Value 'Skip'
New-OfficeExcelValue -Worksheet $sheet -Row 2 -Column 1 -Value 'FirstName'
New-OfficeExcelValue -Worksheet $sheet -Row 2 -Column 2 -Value 'Age'
New-OfficeExcelValue -Worksheet $sheet -Row 3 -Column 1 -Value 'John'
New-OfficeExcelValue -Worksheet $sheet -Row 3 -Column 2 -Value 30
Save-OfficeExcel -Workbook $workbook -FilePath $pathHeader

$rowsWithHeader = Import-OfficeExcel -FilePath $pathHeader -StartRow 2 -EndRow 3 -HeaderRow 2
$rowsWithHeader | Format-Table

$pathNoHeader = "$PSScriptRoot/NoHeader.xlsx"
$workbookNoHeader = New-OfficeExcel
$sheetNoHeader = New-OfficeExcelWorkSheet -Workbook $workbookNoHeader -WorksheetName 'Data' -Option Replace
New-OfficeExcelValue -Worksheet $sheetNoHeader -Row 1 -Column 1 -Value 'John'
New-OfficeExcelValue -Worksheet $sheetNoHeader -Row 1 -Column 2 -Value 30
New-OfficeExcelValue -Worksheet $sheetNoHeader -Row 2 -Column 1 -Value 'Jane'
New-OfficeExcelValue -Worksheet $sheetNoHeader -Row 2 -Column 2 -Value 25
Save-OfficeExcel -Workbook $workbookNoHeader -FilePath $pathNoHeader

$rowsNoHeader = Import-OfficeExcel -FilePath $pathNoHeader -NoHeader -StartRow 1 -EndRow 2 -StartColumn 1 -EndColumn 2
$rowsNoHeader | Format-Table
