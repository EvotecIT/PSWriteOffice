function New-OfficeExcelWorkSheet {
    [cmdletBinding()]
    param(
        [parameter(Position = 0)][scriptblock] $ExcelContent,
        [alias('ExcelDocument')][ClosedXML.Excel.XLWorkbook]$Excel,
        [parameter(Mandatory)][alias('Name')][string] $WorksheetName,
        [ValidateSet("Replace", "Skip", "Rename")][string] $Option = 'Skip',
        [switch] $Suppress,
        [string] $TabColor
    )
    $Worksheet = $null
    # This decides between inline and standalone usage of the command

    if ($null -ne $Excel) {
        # We do nothing
    } elseif ($Script:OfficeTrackerExcel -and -not $Excel) {
        $Excel = $Script:OfficeTrackerExcel['WorkBook']
    } else {
        # Excel not provided, this means most likely some other cmdlet failed up in the chain
        return
    }

    if ($Excel.Worksheets.Contains($WorksheetName)) {
        if ($Option -eq 'Skip') {
            Write-Warning -Message "New-OfficeExcelWorkSheet - WorkSheet with name $WorksheetName already exists. Skipping..."
            return
        } elseif ($Option -eq 'Replace') {
            Write-Warning -Message "New-OfficeExcelWorkSheet - WorkSheet with name $WorksheetName already exists. Replacing..."
            $Excel.Worksheets.Worksheet($WorksheetName).Delete()
            $Worksheet = $Excel.Worksheets.Add($WorksheetName)
        } elseif ($Option -eq 'Rename') {
            $NewName = "Sheet" + (Get-RandomStringName -Size 6)
            Write-Warning -Message "New-OfficeExcelWorkSheet - WorkSheet with name $WorksheetName already exists. Renaming to $NewName..."
            # $Worksheet = $Excel.Worksheets.Worksheet($WorksheetName)
            $WorkSheetName = $NewName
            $Worksheet = $Excel.Worksheets.Add($WorksheetName)
        }
    } else {
        $Worksheet = $Excel.Worksheets.Add($WorksheetName)
    }
    if ($Worksheet) {
        if ($TabColor) {
            Set-OfficeExcelWorkSheetStyle -TabColor $TabColor -Worksheet $Worksheet
        }
        if ($ExcelContent) {
            # This is to support inline mode
            $Script:OfficeTrackerExcel['WorkSheet'] = $Worksheet
            $ExecutedContent = &  $ExcelContent
            $ExecutedContent
            $Script:OfficeTrackerExcel['WorkSheet'] = $null
        } else {
            if (-not $Suppress) {
                # Standalone approach
                if ($NameOnly) {
                    $Worksheet.Name
                } else {
                    $Worksheet
                }
            }
        }
    }
}

Register-ArgumentCompleter -CommandName New-OfficeExcelWorkSheet -ParameterName TabColor -ScriptBlock $Script:ScriptBlockColors