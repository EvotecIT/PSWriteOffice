Import-Module .\PSWriteOffice.psd1 -Force

$ProcessList = Get-Process | Select-Object -First 5
$ProcessList | Export-OfficeExcel -FilePath $PSScriptRoot\Test5.xlsx -WorksheetName 'Contact3' -Show

# cmdlet from PSWriteExcel for speed comparison
#ConvertTo-Excel -FilePath $PSScriptRoot\Test15.xlsx -ExcelWorkSheetName 'Contact3' -DataTable $ProcessList