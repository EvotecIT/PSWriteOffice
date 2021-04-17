function Get-OfficeExcelWorkSheet {
    [cmdletBinding(DefaultParameterSetName = 'All')]
    param(
        [parameter(Mandatory, ParameterSetName = 'Name')]
        [parameter(Mandatory, ParameterSetName = 'Index')]
        [parameter(Mandatory, ParameterSetName = 'All')]
        [alias('ExcelDocument')][ClosedXML.Excel.XLWorkbook]$Excel,

        [parameter(ParameterSetName = 'Name')][alias('Name')][string] $WorksheetName,
        [parameter(ParameterSetName = 'Index')][nullable[int]] $Index,
        [parameter(ParameterSetName = 'All')][switch] $All,

        [parameter(ParameterSetName = 'Name')]
        [parameter(ParameterSetName = 'Index')]
        [parameter(ParameterSetName = 'All')]
        [switch] $NameOnly
    )
    try {
        if ($WorksheetName) {
            $Worksheet = $Excel.Worksheets.Worksheet($WorksheetName)
        } elseif ($null -ne $Index) {
            $Worksheet = $Excel.Worksheets.Worksheet($Index)
        } else {
            $Worksheet = $Excel.Worksheets
        }
        if ($NameOnly) {
            $Worksheet.Name
        } else {
            $Worksheet
        }
    } catch {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning -Message "Get-OfficeExcelWorkSheet - Error: $($_.Exception.Message)"
        }
    }
}