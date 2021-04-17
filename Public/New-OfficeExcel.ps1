function New-OfficeExcel {
    [cmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $FilePath,
        [switch] $Template,
        [nullable[bool]] $RecalculateAllFormulas,
        [ClosedXML.Excel.XLEventTracking] $EventTracking
    )

    if ($FilePath) {
        if (Test-Path -LiteralPath $FilePath) {
            Write-Warning "New-OfficeExcel - File $FilePath already exists. Loading up."
            $WorkBook = [ClosedXML.Excel.XLWorkbook]::new($FilePath)
            $WorkBook | Add-Member -MemberType NoteProperty -Name 'OpenType' -Value 'Existing' -Force
        } else {
            $WorkBook = [ClosedXML.Excel.XLWorkbook]::new()
            $WorkBook | Add-Member -MemberType NoteProperty -Name 'OpenType' -Value 'New' -Force
        }
        $WorkBook | Add-Member -MemberType NoteProperty -Name 'FilePath' -Value $FilePath -Force
        $WorkBook
    }
}