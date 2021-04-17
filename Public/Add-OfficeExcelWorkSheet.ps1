function Add-OfficeExcelWorkSheet {
    [cmdletBinding()]
    param(
        [parameter(Mandatory)][alias('ExcelDocument')][ClosedXML.Excel.XLWorkbook]$Excel,
        [parameter(Mandatory)][alias('Name')][string] $WorksheetName,
        [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Skip',
        [switch] $Suppress
    )

    $WorkBook = $Excel.Worksheets.Add($WorksheetName)
    if (-not $Suppress) {
        $WorkBook
    }
}