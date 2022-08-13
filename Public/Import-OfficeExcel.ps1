function Import-OfficeExcel {
    <#
    .SYNOPSIS
    Provides a way to converting an Excel file into PowerShell objects.

    .DESCRIPTION
    Provides a way to converting an Excel file into PowerShell objects.
    If Worksheet is not specified, all worksheets will be imported and returned as a hashtable of worksheet names and worksheet objects.
    If Worksheet is specified, only the specified worksheet will be imported and returned as an array of PSCustomObjects

    .PARAMETER FilePath
    The path to the Excel file to import.

    .PARAMETER WorkSheetName
    The name of the worksheet to import. If not specified, all worksheets will be imported.

    .EXAMPLE
    $FilePath = "$PSScriptRoot\Documents\Test5.xlsx"

    $ImportedData1 = Import-OfficeExcel -FilePath $FilePath
    $ImportedData1 | Format-Table

    .EXAMPLE
    $FilePath = "$PSScriptRoot\Documents\Excel.xlsx"

    $ImportedData2 = Import-OfficeExcel -FilePath $FilePath -WorkSheetName 'Contact3'
    $ImportedData2 | Format-Table

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Alias('LiteralPath')][Parameter(Mandatory)][string] $FilePath,
        [string[]] $WorkSheetName
    )

    $ExcelWorkbook = Get-OfficeExcel -FilePath $FilePath
    if ($ExcelWorkbook) {
        $WorkSheetContent = [ordered] @{}
        foreach ($WorkSheet in $ExcelWorkbook.Worksheets) {
            # if user asked for specific worksheet we need to deliver
            if ($WorkSheetName) {
                if ($WorkSheet.Worksheet -notin $WorkSheetName) {
                    continue
                }
            }
            $WorkSheetContent[$WorkSheet.Name] = Get-OfficeExcelWorkSheetData -WorkSheet $WorkSheet
        }
        if ($WorkSheetName.Count -eq 1) {
            $WorkSheetContent[$WorkSheetName]
        } elseif ($WorkSheetName.Count -eq 0 -and $WorkSheetContent.Count -eq 1) {
            $WorkSheetContent[0]
        } else {
            $WorkSheetContent
        }
        $ExcelWorkbook.Dispose()
    }
}