function Save-OfficeExcel {
    [cmdletBinding()]
    param(
        [ClosedXML.Excel.XLWorkbook] $Excel,
        [string] $FilePath,
        [switch] $Show
    )
    if ($Excel) {
        if (-not $FilePath) {
            $FilePath = $Excel.FilePath
        }
        if ($Excel.Worksheets.Count -gt 0) {
            if (-not $FilePath) {
                if ($Excel.OpenType -eq 'Existing') {
                    $Excel.Save()
                } else {
                    if ($Excel.OpenType -eq 'New') {
                        $Excel.SaveAs($Excel.FilePath)
                    }
                }
            } else {
                $Excel.SaveAs($FilePath)
            }
            if ($Show) {
                Invoke-Item -Path $FilePath
            }
        } else {
            Write-Warning -Message "Save-OfficeExcel - Can't save $FilePath because there are no worksheets."
        }
    } else {
        Write-Warning -Message "Save-OfficeExcel - Excel Workbook not provided. Skipping."
    }
}