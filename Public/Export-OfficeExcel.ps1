function Export-OfficeExcel {
    [cmdletBinding()]
    param(
        [string] $FilePath,
        [alias('Name')][string] $WorksheetName = 'Sheet1',
        [alias("TargetData")][Parameter(ValueFromPipeline = $true)][Array] $DataTable,
        [int] $Row = 1,
        [int] $Column = 1,
        [switch] $Show,
        [switch] $AllProperties,
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
    Begin {
        $Data = [System.Collections.Generic.List[Object]]::new()
    }
    Process {
        foreach ($_ in $DataTable) {
            $Data.Add($_)
        }
    }
    End {
        New-OfficeExcel -FilePath $FilePath {
            New-OfficeExcelWorkSheet -Name $WorksheetName {
                $SplatOfficeExcelTable = @{
                    DataTable            = $Data
                    Row                  = $Row
                    Column               = $Column
                    AllProperties        = $AllProperties.IsPresent
                    DisableAutoFilter    = $DisableAutoFilter.IsPresent
                    EmphasizeFirstColumn = $EmphasizeFirstColumn.IsPresent
                    EmphasizeLastColumn  = $EmphasizeLastColumn.IsPresent
                    ShowColumnStripes    = $ShowColumnStripes.IsPresent
                    ShowRowStripes       = $ShowRowStripes.IsPresent
                    ShowTotalsRow        = $ShowTotalsRow.IsPresent
                    HideHeaderRow        = $HideHeaderRow.IsPresent
                    Transpose            = $Transpose
                    Theme                = $Theme
                }
                Remove-EmptyValue -Hashtable $SplatOfficeExcelTable
                New-OfficeExcelTable @SplatOfficeExcelTable #-DataTable $Data -Row $Row -Column $Column -AllProperties:$AllProperties -AutoFilter -Transpose $Transpose
            } -Option Replace
        } -Show:$Show.IsPresent -Save
    }
}

<#
$Script:ScriptBlockThemes = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    [ClosedXML.Excel.XLTableTheme]::GetAllThemes() | Where-Object { $_ -like "*$wordToComplete*" }
}
#>

Register-ArgumentCompleter -CommandName Export-OfficeExcel -ParameterName Theme -ScriptBlock $Script:ScriptBlockThemes