function Get-OfficeExcelValue {
    [cmdletBinding()]
    param(
        [ClosedXML.Excel.IXLWorksheet] $Worksheet,
        [int] $Row,
        [int] $Column
    )
    if ($Script:OfficeTrackerExcel) {
        $Worksheet = $Script:OfficeTrackerExcel['WorkSheet']
    } elseif (-not $Worksheet) {
        return
    }

    $Worksheet.Cell($Row, $Column)
}