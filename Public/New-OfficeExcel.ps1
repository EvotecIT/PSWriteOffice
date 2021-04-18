function New-OfficeExcel {
    [cmdletBinding()]
    param(
        [scriptblock] $ExcelContent,
        [Parameter(Mandatory)][string] $FilePath,
        [switch] $Template,
        [nullable[bool]] $RecalculateAllFormulas,
        [ClosedXML.Excel.XLEventTracking] $EventTracking,
        [switch] $Show,
        [switch] $Save,
        [validateSet('Reuse', 'Overwrite', 'Stop')][string] $WhenExists = 'Reuse'
    )
    if ($ExcelContent) {
        $Script:OfficeTrackerExcel = [ordered] @{}
    }

    if (Test-Path -LiteralPath $FilePath) {
        if ($WhenExists -eq 'Stop') {
            Write-Warning -Message "New-OfficeExcel - File $FilePath already exists. Terminating."
            # lets clean up
            Remove-Variable -Name $Script:OfficeTrackerExcel
            return
        } elseif ($WhenExists -eq 'Overwrite') {
            $WorkBook = [ClosedXML.Excel.XLWorkbook]::new()
            $WorkBook | Add-Member -MemberType NoteProperty -Name 'OpenType' -Value 'New' -Force
        } elseif ($WhenExists -eq 'ReUse') {
            Write-Warning -Message "New-OfficeExcel - File $FilePath already exists. Loading up."
            try {
                $WorkBook = [ClosedXML.Excel.XLWorkbook]::new($FilePath)
            } catch {
                # lets clean up
                Remove-Variable -Name $Script:OfficeTrackerExcel
                if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                    throw
                } else {
                    Write-Warning -Message "New-OfficeExcel - File $FilePath returned error: $($_.Exception.Message)"
                    return
                }
            }
            $WorkBook | Add-Member -MemberType NoteProperty -Name 'OpenType' -Value 'Existing' -Force
        }
    } else {
        $WorkBook = [ClosedXML.Excel.XLWorkbook]::new()
        $WorkBook | Add-Member -MemberType NoteProperty -Name 'OpenType' -Value 'New' -Force
    }
    $WorkBook | Add-Member -MemberType NoteProperty -Name 'FilePath' -Value $FilePath -Force

    # Lets execute what user wanted to execute
    if ($ExcelContent) {
        $Script:OfficeTrackerExcel['WorkBook'] = $WorkBook
        $Script:OfficeTrackerExcel['OpenType'] = 'Existing'
        $ExecutedContent = & $ExcelContent
        $ExecutedContent
    }

    # This means we use all in one cmdlet, so we're saving
    if ($ExcelContent) {
        if ($Save) {
            Save-OfficeExcel -Show:$Show.IsPresent-FilePath $FilePath -Excel $WorkBook
        }
        # lets clean up
        $Script:OfficeTrackerExcel = $null
    } else {
        $WorkBook
    }
}