Import-Module .\PSWriteOffice.psd1 -Force

$ProcessList = Get-Process | Select-Object -First 50
Export-OfficeExcel -FilePath $PSScriptRoot\Documents\Test5.xlsx -WorksheetName 'Contact3' -DataTable $ProcessList  #-Show

# cmdlet from PSWriteExcel for speed comparison
# ConvertTo-Excel -FilePath $PSScriptRoot\Documents\Test15.xlsx -ExcelWorkSheetName 'Contact3' -DataTable $ProcessList #-OpenWorkBook