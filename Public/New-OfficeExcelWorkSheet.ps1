function New-OfficeExcelWorkSheet {
    [cmdletBinding()]
    param(
        [parameter(Position = 0)][scriptblock] $ExcelContent,
        [alias('ExcelDocument')][ClosedXML.Excel.XLWorkbook]$Excel,
        [parameter(Mandatory)][alias('Name')][string] $WorksheetName,
        [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Skip',
        [switch] $Suppress
    )
    $Worksheet = $null
    # This decides between inline and standalone usage of the command
    if ($Script:OfficeTrackerExcel -and -not $Excel) {
        $Excel = $Script:OfficeTrackerExcel['WorkBook']
    }
    if ($Excel.Worksheets.Contains($WorksheetName)) {
        $Worksheet = $Excel.Worksheets.Worksheet($WorksheetName)
    } else {
        $Worksheet = $Excel.Worksheets.Add($WorksheetName)
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
        if (-not $Suppress) {
            $Worksheet
        }
    } else {

    }
}