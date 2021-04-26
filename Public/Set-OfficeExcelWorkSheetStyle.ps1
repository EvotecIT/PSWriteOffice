function Set-OfficeExcelWorkSheetStyle {
    [cmdletBinding(DefaultParameterSetName = 'Name')]
    param(
        [parameter(ParameterSetName = 'Name')]
        [parameter(ParameterSetName = 'Index')]
        [parameter(ParameterSetName = 'Native')]
        [alias('ExcelDocument')][ClosedXML.Excel.XLWorkbook]$Excel,
        [parameter(ParameterSetName = 'Name')]
        [parameter(ParameterSetName = 'Index')]
        [parameter(ParameterSetName = 'Native')]
        [string] $TabColor,
        [parameter(ParameterSetName = 'Native')] $Worksheet,
        [parameter(ParameterSetName = 'Name')][alias('Name')][string] $WorksheetName,
        [parameter(ParameterSetName = 'Index')][nullable[int]] $Index
    )
    #$Worksheet = $null
    # This decides between inline and standalone usage of the command
    if ($Script:OfficeTrackerExcel -and -not $Excel) {
        $Excel = $Script:OfficeTrackerExcel['WorkBook']
    }
    # Lets get worksheet we need
    if ($Worksheet) {
        # this means we provided worksheet object
    } else {
        try {
            if ($WorksheetName) {
                $Worksheet = $Excel.Worksheets.Worksheet($WorksheetName)
            } elseif ($null -ne $Index) {
                $Worksheet = $Excel.Worksheets.Worksheet($Index)
            }
        } catch {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw
            } else {
                Write-Warning -Message "Set-OfficeExcelWorkSheet - Error: $($_.Exception.Message)"
            }
        }
    }
    if ($Worksheet) {
        if ($TabColor) {
            $ColorConverted = [ClosedXML.Excel.XLColor]::FromHtml((ConvertFrom-Color -Color $TabColor))
            $null = $Worksheet.SetTabColor($ColorConverted)
        }
    }
}

Register-ArgumentCompleter -CommandName Set-OfficeExcelWorkSheetStyle -ParameterName TabColor -ScriptBlock $Script:ScriptBlockColors