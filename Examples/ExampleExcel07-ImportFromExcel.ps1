Import-Module .\PSWriteOffice.psd1 -Force

$FilePath = "$PSScriptRoot\Documents\Test5.xlsx"

$ImportedData1 = Import-OfficeExcel -FilePath $FilePath
$ImportedData1 | Format-Table

$FilePath = "$PSScriptRoot\Documents\Excel.xlsx"

$ImportedData2 = Import-OfficeExcel -FilePath $FilePath -WorkSheetName 'Contact3'
$ImportedData2 | Format-Table