function Save-OfficeExcel {
    [cmdletBinding()]
    param(
        [ClosedXML.Excel.XLWorkbook] $Excel,
        [string] $FilePath,
        [switch] $Show
    )

    if (-not $FilePath) {
        if ($Excel.OpenType -eq 'Existing') {
            $Excel.Save()
        } else {
            if ($Excel.OpenType -eq 'New') {
                $Excel.SaveAs($Excel.FilePath)
            }
        }
        $FilePath = $Excel.FilePath
    } else {
        $Excel.SaveAs($FilePath)
    }
    if ($Show) {
        Invoke-Item -Path $FilePath
    }
}