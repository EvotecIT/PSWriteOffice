function New-OfficeExcelValue {
    [cmdletBinding()]
    param(
        $Worksheet,
        [Object] $Value,
        [int] $Row,
        [int] $Column
    )

    if ($Script:OfficeTrackerExcel) {
        $Worksheet = $Script:OfficeTrackerExcel['WorkSheet']
    } elseif (-not $Worksheet) {
        return
    }

    $Worksheet.Cell($Row, $Column).Value = $Value
}