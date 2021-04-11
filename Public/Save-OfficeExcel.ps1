function Save-OfficeExcel {
    [cmdletBinding()]
    param(
        $Excel,
        [string] $FilePath,
        [switch] $Show
    )

    if (-not $FilePath) {
        $Excel.Save()
        $FilePath = $Excel.FilePath
    } else {
        $Excel.SaveAs($FilePath)

    }
    if ($Show) {
        Invoke-Item -Path $FilePath
    }
}