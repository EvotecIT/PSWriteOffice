Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

#$Excel = New-OfficeExcel -FilePath $PSScriptRoot\Documents\Excel.xlsx
#$Excel = Get-OfficeExcel -FilePath "C:\Support\GitHub\PSWriteOffice\Examples\Documents\Excel1.xlsx"
$Excel = New-OfficeExcel -FilePath "C:\Support\GitHub\PSWriteOffice\Examples\Documents\Excel1.xlsx"

#Add-OfficeExcelWorkSheet -Excel $Excel -WorksheetName 'Contact1' -Suppress
#Add-OfficeExcelWorkSheet -Excel $Excel -WorksheetName 'Contact2' -Suppress

$Worksheet = Get-OfficeExcelWorkSheet -Excel $Excel -Name 'Contact1'
#Get-OfficeExcelWorkSheet -Excel $Excel -Index 2 -NameOnly | Format-Table

#$Worksheet = Add-OfficeExcelWorkSheet -Excel $Excel -WorksheetName 'Contact3'

$Strings = @(
    'Test'
    'Test2'
    'Test3'
)
$Objects = @(
    [PSCustomObject] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
    [PSCustomObject] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
)

$Objects2 = @(
    [ordered] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
    [ordered] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
)



#$Worksheet.Cell(1, 1).Value = 'FromStrings'
#$Worksheet.Range(1, 3, 1, 8).Merge().AddToNamed('Titles')
#$Worksheet.Cell(2, 1).InsertData($Strings)
#$Worksheet.cell(10, 1).InsertTable($table)



#Add-OfficeExcelWorkSheet -Excel $Excel
#$WorkSheet = $Excel.Worksheets.Add('Contacts3')

<#
$WorkSheet = $Excel.Worksheets | Where-Object { $_.Name -eq 'Contacts3' }
$WorkSheet.Cell("A1").Value = "Hello World!";
$WorkSheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";

#>


#$WorkSheet = $Excel.Worksheets | Where-Object { $_.Name -eq 'Sheet1' }


New-OfficeExcelTable -DataTable $Objects -Worksheet $Worksheet -Row 30 -Column 5
New-OfficeExcelTable -DataTable $Objects2 -Worksheet $Worksheet -Row 40 -Column 5

Save-OfficeExcel -Excel $Excel -Show