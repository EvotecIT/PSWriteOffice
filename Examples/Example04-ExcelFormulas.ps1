Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

#$Excel = New-OfficeExcel -FilePath $PSScriptRoot\Documents\Excel.xlsx
$Excel = New-OfficeExcel -FilePath "C:\Support\GitHub\PSWriteOffice\Examples\Documents\pswriteexcel_cell.xlsx"

#$WorkSheet = $Excel.Worksheets.Add('Contacts3')

<#
$WorkSheet = $Excel.Worksheets | Where-Object { $_.Name -eq 'Contacts3' }
$WorkSheet.Cell("A1").Value = "Hello World!";
$WorkSheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";

#>


$WorkSheet = $Excel.Worksheets | Where-Object { $_.Name -eq 'Sheet1' }


foreach ($Cell in $WorkSheet.Cells()) {
    if ($Cell.FormulaA1) {
        if ($Cell.FormulaA1.StartsWith('_xlfn.FORMULATEXT')) {
            $Cell.FormulaA1 = $Cell.CachedValue
        }
    }
}

$WorkSheet.RecalculateAllFormulas()

#$WorkSheet.Cell('E1') | Format-Table *
#$WorkSheet.Cell('E1').FormulaA1 = $WorkSheet.Cell('E1').CachedValue

$WorkSheet.cell('C1').Value = 30.5
$WorkSheet.RecalculateAllFormulas()

$WorkSheet.CellsUsed('E') | Format-Table
#Save-OfficeExcel -Excel $Excel -Show