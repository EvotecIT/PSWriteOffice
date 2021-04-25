function New-OfficeExcelValue {
    [cmdletBinding()]
    param(
        $Worksheet,
        [Object] $Value,
        [int] $Row,
        [int] $Column,
        [string] $DateFormat,
        [string] $NumberFormat,
        [int] $FormatID
    )
    $KnownTypes = 'bool|byte|char|datetime|decimal|double|float|int|long|sbyte|short|string|timespan|uint|ulong|URI|ushort'
    if ($Script:OfficeTrackerExcel) {
        $Worksheet = $Script:OfficeTrackerExcel['WorkSheet']
    } elseif (-not $Worksheet) {
        return
    }

    try {
        if ($null -eq $Value) {
            $Worksheet.Cell($Row, $Column).Value = ''
        } elseif ($Value.GetType().Name -match $KnownTypes) {
            $Worksheet.Cell($Row, $Column).Value = $Value
        } else {
            $Worksheet.Cell($Row, $Column).Value = [string] $Value
        }
    } catch {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning "New-OfficeExcelValue - Error: $($_.Exception.Message)"
        }
    }
}