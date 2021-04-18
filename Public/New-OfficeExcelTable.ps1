function New-OfficeExcelTable {
    [cmdletBinding()]
    param(
        [Array] $DataTable,
        [Parameter(Mandatory)][int] $Row,
        [Parameter(Mandatory)][int] $Column,
        $Worksheet,
        [switch] $AllProperties,
        [switch] $ReturnObject,
        [ClosedXML.Excel.XLTransposeOptions] $Transpose,
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

    [System.Data.DataTable] $Table = [System.Data.DataTable]::new('MyTable')
    if ($DataTable.Count -eq 0) {
        return
    }

    if ($DataTable[0] -isnot [System.Collections.IDictionary]) {
        $Properties = Select-Properties -Objects $DataTable -AllProperties:$AllProperties -Property $IncludeProperty -ExcludeProperty $ExcludeProperty
    } else {
        $Properties = 'Name', 'Value'
    }

    foreach ($Property in $Properties) {
        $null = $table.Columns.Add($Property, [string])
    }

    # Add data
    foreach ($Object in $DataTable) {
        if ($Object -isnot [System.Collections.IDictionary]) {
            $Values = foreach ($Property in $Properties) {
                $Value = $Object.$Property
                if ($Value.Count -gt 1) {
                    $Value = $Value -join ','
                }
                $Value
            }
            try {
                $null = $table.Rows.Add($Values)
            } catch {
                Write-Warning -Message "New-OfficeExcelTable - Error when adding values to row $($_.Exception.Message). Skipping."
            }
        } else {
            foreach ($Entry in $Object.Keys) {
                $null = $table.Rows.Add($Entry, $Object[$Entry])
            }
        }
    }

    try {
        $TableOutput = $Worksheet.cell($Row, $Column).InsertTable($table)
    } catch {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning -Message "New-OfficeExcelTable - Error occured: $($_.Exception.Message)"
        }
    }
    if ($null -ne $Transpose) {
        $TableOutput.Transpose($Transpose)
    }
    if ($AutoFilter) {
        $TableOutput.InitializeAutoFilter()
    }
    if ($ShowColumnStripes) {
        $TableOutput.ShowColumnStripes = $true
    }
    if ($ShowRowStripes) {
        $TableOutput.ShowRowStripes = $true
    }
    if ($DisableAutoFilter) {
        $TableOutput.ShowAutoFilter = $false
    }
    if ($ShowTotalsRow) {
        $TableOutput.ShowsTotalRow = $true
    }
    if ($null -ne $Theme) {
        $TableOutput.Theme = $Theme
    }
    if ($EmphasizeFirstColumn) {
        $TableOutput.EmphasizeFirstColumn = $true
    }
    if ($EmphasizeLastColumn) {
        $TableOutput.EmphasizeLastColumn = $true
    }
    if ($HideHeaderRow) {
        $TableOutput.ShowHeaderRow = $false
    }
    if ($ReturnObject) {
        $TableOutput
    }
}

$Script:ScriptBlockThemes = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    [ClosedXML.Excel.XLTableTheme]::GetAllThemes() | Where-Object { $_ -like "*$wordToComplete*" }
}

Register-ArgumentCompleter -CommandName New-OfficeExcelTable -ParameterName Theme -ScriptBlock $Script:ScriptBlockThemes