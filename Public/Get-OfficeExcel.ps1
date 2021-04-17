function Get-OfficeExcel {
    [cmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $FilePath,
        [switch] $Template,
        [nullable[bool]] $RecalculateAllFormulas,
        [ClosedXML.Excel.XLEventTracking] $EventTracking
    )

    if ($FilePath -and (Test-Path -LiteralPath $FilePath)) {
        if ($RecalculateAllFormulas -or $EventTracking) {
            $LoadOptions = [ClosedXML.Excel.LoadOptions]::new()
            if ($null -ne $RecalculateAllFormulas) {
                $LoadOptions.RecalculateAllFormulas = $RecalculateAllFormulas
            }
            if ($EventTracking) {
                $LoadOptions.EventTracking = $EventTracking
            }
            $WorkBook = [ClosedXML.Excel.XLWorkbook]::new($FilePath, $LoadOptions)
        } else {
            if ($FilePath) {
                $WorkBook = [ClosedXML.Excel.XLWorkbook]::new($FilePath)
            } else {
                $WorkBook = [ClosedXML.Excel.XLWorkbook]::new()
            }
        }
        $WorkBook | Add-Member -MemberType NoteProperty -Name 'FilePath' -Value $FilePath -Force
        $WorkBook
    }
    <#
    ClosedXML.Excel.XLWorkbook new()
    ClosedXML.Excel.XLWorkbook new(ClosedXML.Excel.XLEventTracking eventTracking)
    ClosedXML.Excel.XLWorkbook new(ClosedXML.Excel.LoadOptions loadOptions)
    ClosedXML.Excel.XLWorkbook new(string file)
    ClosedXML.Excel.XLWorkbook new(string file, ClosedXML.Excel.XLEventTracking eventTracking)
    ClosedXML.Excel.XLWorkbook new(string file, ClosedXML.Excel.LoadOptions loadOptions)
    ClosedXML.Excel.XLWorkbook new(System.IO.Stream stream)
    ClosedXML.Excel.XLWorkbook new(System.IO.Stream stream, ClosedXML.Excel.XLEventTracking eventTracking)
    ClosedXML.Excel.XLWorkbook new(System.IO.Stream stream, ClosedXML.Excel.LoadOptions loadOptions)
    #>
}