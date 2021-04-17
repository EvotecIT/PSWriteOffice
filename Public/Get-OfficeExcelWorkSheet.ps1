function Get-OfficeExcelWorkSheet {
    [cmdletBinding(DefaultParameterSetName = 'All')]
    param(
        [parameter(Position = 0, ParameterSetName = 'Name')]
        [parameter(Position = 0, ParameterSetName = 'Index')]
        [parameter(Position = 0, ParameterSetName = 'All')]
        [scriptblock] $ExcelContent,

        [parameter(ParameterSetName = 'Name')]
        [parameter(ParameterSetName = 'Index')]
        [parameter(ParameterSetName = 'All')]
        [alias('ExcelDocument')][ClosedXML.Excel.XLWorkbook]$Excel,

        [parameter(ParameterSetName = 'Name')][alias('Name')][string] $WorksheetName,
        [parameter(ParameterSetName = 'Index')][nullable[int]] $Index,
        [parameter(ParameterSetName = 'All')][switch] $All,

        [parameter(ParameterSetName = 'Name')]
        [parameter(ParameterSetName = 'Index')]
        [parameter(ParameterSetName = 'All')]
        [switch] $NameOnly
    )
    $Worksheet = $null
    # This decides between inline and standalone usage of the command
    if ($Script:OfficeTrackerExcel -and -not $Excel) {
        $Excel = $Script:OfficeTrackerExcel['WorkBook']
    }
    try {
        if ($WorksheetName) {
            $Worksheet = $Excel.Worksheets.Worksheet($WorksheetName)
        } elseif ($null -ne $Index) {
            $Worksheet = $Excel.Worksheets.Worksheet($Index)
        } else {
            $Worksheet = $Excel.Worksheets
        }
    } catch {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning -Message "Get-OfficeExcelWorkSheet - Error: $($_.Exception.Message)"
        }
    }
    if ($Worksheet) {
        if ($ExcelContent) {
            # This is to support inline mode
            $Script:OfficeTrackerExcel['WorkSheet'] = $Worksheet
            $ExecutedContent = &  $ExcelContent
            $ExecutedContent
            $Script:OfficeTrackerExcel['WorkSheet'] = $null
        } else {
            # Standalone approach
            if ($NameOnly) {
                $Worksheet.Name
            } else {
                $Worksheet
            }
        }
    } else {
        if ($Index) {
            Write-Warning -Message "Get-OfficeExcelWorkSheet - WorkSheet with index $Index doesnt exits. Skipping."
        } elseif ($WorksheetName) {
            Write-Warning -Message "Get-OfficeExcelWorkSheet - WorkSheet with name $WorksheetName doesnt exits. Skipping."
        } else {
            Write-Warning -Message "Get-OfficeExcelWorkSheet - Mmm?"
        }
    }
}