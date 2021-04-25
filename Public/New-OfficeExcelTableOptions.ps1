function New-OfficeExcelTableOptions {
    [cmdletBinding()]
    param(
        $Table,
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

    # Apply some options to table we just added
    if ($Table) {
        if ($null -ne $Transpose) {
            $Table.Transpose($Transpose)
        }
        if ($AutoFilter) {
            $Table.InitializeAutoFilter()
        }
        if ($ShowColumnStripes) {
            $Table.ShowColumnStripes = $true
        }
        if ($ShowRowStripes) {
            $Table.ShowRowStripes = $true
        }
        if ($DisableAutoFilter) {
            $Table.ShowAutoFilter = $false
        }
        if ($ShowTotalsRow) {
            $Table.ShowsTotalRow = $true
        }
        if ($null -ne $Theme) {
            $Table.Theme = $Theme
        }
        if ($EmphasizeFirstColumn) {
            $Table.EmphasizeFirstColumn = $true
        }
        if ($EmphasizeLastColumn) {
            $Table.EmphasizeLastColumn = $true
        }
        if ($HideHeaderRow) {
            $Table.ShowHeaderRow = $false
        }
        if ($ReturnObject) {
            $Table
        }
    }
}