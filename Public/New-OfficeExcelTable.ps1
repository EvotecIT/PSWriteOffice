function New-OfficeExcelTable {
    [cmdletBinding()]
    param(
        [Array] $DataTable,
        [Object] $Worksheet,
        [alias('Row')][int] $StartRow = 1,
        [alias('Column')][int] $StartCell = 1,
        [switch] $ReturnObject,
        [ClosedXML.Excel.XLTransposeOptions] $Transpose,
        [switch] $AllProperties,
        [switch] $SkipHeader,
        [switch] $ShowRowStripes,
        [switch] $ShowColumnStripes,
        [switch] $DisableAutoFilter,
        [switch] $HideHeaderRow,
        [switch] $ShowTotalsRow,
        [switch] $EmphasizeFirstColumn,
        [switch] $EmphasizeLastColumn,
        [string] $Theme
    )
    # This decides between inline and standalone usage of the command
    if ($Script:OfficeTrackerExcel -and -not $Worksheet) {
        $WorkSheet = $Script:OfficeTrackerExcel['WorkSheet']
    }

    $Cell = $StartCell - 1
    # Table header
    if ($DataTable[0] -is [System.Collections.IDictionary]) {
        $Properties = 'Name', 'Value'
    } else {
        $Properties = Select-Properties -Objects $DataTable -AllProperties:$AllProperties -Property $IncludeProperty -ExcludeProperty $ExcludeProperty
    }
    # Add Table Header (Title)
    if (-not $SkipHeader) {
        foreach ($Property in $Properties) {
            $Cell++
            New-OfficeExcelValue -Row $StartRow -Value $Property -Column $Cell -Worksheet $Worksheet
        }
    }
    # Table content
    if ($DataTable[0] -is [System.Collections.IDictionary]) {
        # By Ordered Dictionary
        #$Row = 1 # we already added header
        foreach ($Data in $DataTable) {
            foreach ($Key in $Data.Keys) {
                $Row++
                New-OfficeExcelValue -Row ($Row + $StartRow) -Value $Key -Column ($StartCell) -Worksheet $Worksheet
                New-OfficeExcelValue -Row ($Row + $StartRow) -Value $Data[$Key] -Column ($StartCell + 1) -Worksheet $Worksheet
            }
        }
        $LastCell = $Worksheet.Row($StartRow + $Row).Cell($Cell)
    } elseif ($Properties -eq '*') {
        foreach ($Data in $DataTable) {
            $Row++
            New-OfficeExcelValue -Row ($Row + $StartRow) -Value $Data -Column ($StartCell) -Worksheet $Worksheet
        }
        $LastCell = $Worksheet.Row($StartRow + $Row).Cell($StartCell)
    } else {
        # By PSCustomObject
        for ($Row = 1; $Row -le $DataTable.Count; $Row++) {
            $Cell = $StartCell - 1
            foreach ($Property in $Properties) {
                $Cell++
                New-OfficeExcelValue -Row ($Row + $StartRow) -Value $DataTable[$Row - 1].$Property -Column $Cell -Worksheet $Worksheet
            }
        }
        $LastCell = $Worksheet.Row($StartRow - 1 + $Row).Cell($Cell)
    }
    $FirstCell = $Worksheet.Row($StartRow).Cell($StartCell)
    $Range = $Worksheet.Range($FirstCell.Address, $LastCell.Address)
    $TableOutput = $Range.CreateTable()

    $SplatOptions = @{
        Table                = $TableOutput
        Transpose            = $Transpose
        ShowRowStripes       = $ShowRowStripes.IsPresent
        ShowColumnStripes    = $ShowColumnStripes.IsPresent
        DisableAutoFilter    = $DisableAutoFilter.IsPresent
        HideHeaderRow        = $HideHeaderRow.IsPresent
        ShowTotalsRow        = $ShowTotalsRow.IsPresent
        EmphasizeFirstColumn = $EmphasizeFirstColumn.IsPresent
        EmphasizeLastColumn  = $EmphasizeLastColumn.IsPresent
        Theme                = $Theme
    }
    Remove-EmptyValue -Hashtable $SplatOptions
    New-OfficeExcelTableOptions @SplatOptions
}