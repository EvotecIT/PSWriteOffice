Import-Module .\PSWriteOffice.psd1 -Force

$ProcessList = Get-Process | Select-Object -First 5
Export-OfficeExcel -FilePath $PSScriptRoot\Documents\Test5.xlsx -WorksheetName 'Contact3' -DataTable $ProcessList -Show