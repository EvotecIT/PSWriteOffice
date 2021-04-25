function New-OfficeExcelTableOld {
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
        $Properties = Select-Properties -Objects $DataTable -AllProperties:$AllProperties -Property $IncludeProperty -ExcludeProperty $ExcludeProperty -IncludeTypes
    } else {
        $Properties = 'Name', 'Value'
    }

    if ($Properties.Name -eq '*') {
        try {
            $TableOutput = $Worksheet.cell($Row, $Column).InsertTable($DataTable)
        } catch {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw
            } else {
                Write-Warning -Message "New-OfficeExcelTable - Error occured: $($_.Exception.Message)"
            }
        }
    } else {
        #foreach ($Property in $Properties) {
        #    $null = $table.Columns.Add($Property, [string])
        #}

        # foreach ($Property in $Properties.Keys) {

        #$null = $table.Columns.Add($Property.Name, $Property.Type)
        #}

        for ($i = 0; $i -lt $Properties['Name'].Count; $i++) {
            $KnownTypes = 'bool|byte|char|datetime|decimal|double|float|int|long|sbyte|short|string|timespan|uint|ulong|URI|ushort'
            if ($Properties['Type'][$i] -match $KnownTypes) {
                # $null = $table.Columns.Add($Properties['Name'][$i], $Properties['Type'][$i])
                $null = $table.Columns.Add($Properties['Name'][$i], [string])
            } else {
                $null = $table.Columns.Add($Properties['Name'][$i], [string])
            }
        }

        # Add data
        foreach ($Object in $DataTable) {
            if ($Object -is [System.Collections.IDictionary]) {
                foreach ($Entry in $Object.Keys) {
                    $null = $table.Rows.Add($Entry, $Object[$Entry])
                }
            } elseif ($Object.GetType().Name -match 'bool|byte|char|datetime|decimal|double|ExcelHyperLink|float|int|long|sbyte|short|string|timespan|uint|ulong|URI|ushort') {

            } else {
                #      if ($Object -isnot [System.Collections.IDictionary]) {
                $Values = foreach ($Property in $Properties.Name) {
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
            }
        }
        try {
            $TableOutput = $Worksheet.cell($Row, $Column).InsertTable($Table)
        } catch {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw
            } else {
                Write-Warning -Message "New-OfficeExcelTable - Error occured: $($_.Exception.Message)"
            }
        }
    }
    # Apply some options to table we just added

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

    <#
    if ($TableOutput) {
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
    #>
}

$Script:ScriptBlockThemes = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    [ClosedXML.Excel.XLTableTheme]::GetAllThemes() | Where-Object { $_ -like "*$wordToComplete*" }
}

Register-ArgumentCompleter -CommandName New-OfficeExcelTable -ParameterName Theme -ScriptBlock $Script:ScriptBlockThemes