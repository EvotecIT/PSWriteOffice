function New-OfficeExcelTable {
    [cmdletBinding()]
    param(
        [Array] $DataTable,
        $Worksheet,
        [switch] $AllProperties,
        [int] $Row,
        [int] $Column
    )

    [System.Data.DataTable] $table = [System.Data.DataTable]::new('MyTable')
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
            $null = $table.Rows.Add(($Object | Select-Object -Property $Properties).PSobject.properties.value)
        } else {
            foreach ($Entry in $Object.Keys) {
                $null = $table.Rows.Add($Entry, $Object[$Entry])
            }
        }
    }
    try {
        $Worksheet.cell($Row, $Column).InsertTable($table)
    } catch {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning -Message "New-OfficeExcelTable - Error occured: $($_.Exception.Message)"
        }
    }
}