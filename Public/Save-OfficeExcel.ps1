function Save-OfficeExcel {
    [cmdletBinding()]
    param(
        $Excel,
        [switch] $Show
    )
    $Excel.Workbook.Save()
    $Excel.Close()
}