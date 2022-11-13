function Get-OfficeExcelWorkSheetData {
    [cmdletBinding()]
    param(
        [ClosedXML.Excel.IXLWorksheet] $WorkSheet
    )

    $HeaderNames = [System.Collections.Generic.List[string]]::new()
    foreach ($Cell in $WorkSheet.RangeUsed().Row(1).Cells()) {
        if ($Cell.InnerText -ne "") {
            $Name = $Cell.InnerText
        } else {
            $Name = "NoName$($Cell.Address)"
        }
        # We need to check for header duplicates, if someone made a mistake
        if ($HeaderNames.Contains($Name)) {
            $Name = $Name + $($Cell.Address)
        }
        $HeaderNames.Add($Name)
    }
    $LastRowUsed = $WorkSheet.RangeUsed().RowCount()

    foreach ($Row in $WorkSheet.RangeUsed().Rows(2, $LastRowUsed)) {
        $RowData = [ordered] @{}
        for ($i = 0; $i -lt $HeaderNames.Count; $i++) {
            $RowData[$HeaderNames[$i]] = $Row.Cells($i + 1).CachedValue
        }
        [PSCustomObject] $RowData
    }
}