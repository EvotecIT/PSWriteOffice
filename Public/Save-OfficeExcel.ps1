function Save-OfficeExcel {
    [cmdletBinding()]
    param(
        [ClosedXML.Excel.XLWorkbook] $Excel,
        [string] $FilePath,
        [switch] $Show,
        [int] $RetryCount = 1,
        [Parameter(DontShow)] $CurrentRetryCount
    )
    if ($Excel) {
        if (-not $FilePath) {
            $FilePath = $Excel.FilePath
        }
        if ($Excel.Worksheets.Count -gt 0) {
            try {
                if (-not $FilePath) {
                    if ($Excel.OpenType -eq 'Existing') {
                        $Excel.Save()
                    } else {
                        if ($Excel.OpenType -eq 'New') {
                            $Excel.SaveAs($Excel.FilePath)
                        }
                    }
                } else {
                    $Excel.SaveAs($FilePath)
                }
                $CurrentRetryCount = 0
            } catch {
                if ($RetryCount -eq $CurrentRetryCount) {
                    Write-Warning "Save-ExcelDocument - Couldnt save Excel to $FilePath. Retry count limit reached. Terminating.."
                    return
                }
                $CurrentRetryCount++
                $ErrorMessage = $_.Exception.Message
                if ($ErrorMessage -like "*The process cannot access the file*because it is being used by another process.*" -or
                    $ErrorMessage -like "*Error saving file*") {
                    $FilePath = Get-FileName -Temporary -Extension 'xlsx'
                    Write-Warning "Save-OfficeExcel - Couldn't save file as it was in use or otherwise. Trying different name $FilePath"
                    Save-OfficeExcel -Excel $Excel -Show:$Show -FilePath $FilePath -RetryCount $RetryCount -CurrentRetryCount $CurrentRetryCount
                    # we return as we already show it within nested Save-OfficeExcel
                    # otherwise we would end up opening things again
                    return
                } else {
                    Write-Warning "Save-OfficeExcel - Error: $ErrorMessage"
                }
            }

            if ($Show) {
                try {
                    Invoke-Item -Path $FilePath
                } catch {
                    Write-Warning "Save-OfficeExcel - Couldn't open file $FilePath as requested."
                }
            }
        } else {
            Write-Warning -Message "Save-OfficeExcel - Can't save $FilePath because there are no worksheets."
        }
    } else {
        Write-Warning -Message "Save-OfficeExcel - Excel Workbook not provided. Skipping."
    }
}