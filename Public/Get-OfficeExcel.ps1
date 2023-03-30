function Get-OfficeExcel {
    [cmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $FilePath,
        [switch] $Template,
        [nullable[bool]] $RecalculateAllFormulas #,
       # [ClosedXML.Excel.XLEventTracking] $EventTracking
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
                try {
                    $WorkBook = [ClosedXML.Excel.XLWorkbook]::new($FilePath)
                } catch {
                    if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                        throw
                    } else {
                        Write-Warning -Message "Get-OfficeExcel - Failed to open $FilePath. Eror: $($_.Exception.Message)"
                        return
                    }
                }
            } else {
                $WorkBook = [ClosedXML.Excel.XLWorkbook]::new()
            }
        }
        $WorkBook | Add-Member -MemberType NoteProperty -Name 'FilePath' -Value $FilePath -Force
        $WorkBook
    }
}