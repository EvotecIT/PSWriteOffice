function Save-OfficeWord {
    [cmdletBinding()]
    param(
        [alias('WordDocument')][OfficeIMO.Word.WordDocument] $Document,
        [switch] $Show,
        [string] $FilePath,
        [int] $Retry = 2
    )
    if (-not $Document) {
        Write-Warning "Save-OfficeWord - Couldn't save Word Document. Document is null."
        return
    }
    if (-not $Document.FilePath -and -not $FilePath) {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning "Save-OfficeWord - Couldn't save Word Document. No file path provided."
            return
        }
    }
    if ($FilePath) {
        # File path was given so we use SaveAs
        try {
            $null = $Document.Save($FilePath, $Show.IsPresent)
        } catch {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw
            } else {
                Write-Warning "Save-OfficeWord - Couldn't save $FilePath. Error: $($_.Exception.Message)"
            }
        } finally {
            #$NewDocument.Dispose()
            #$Document.Dispose()
        }
    } else {
        if (-not $Document.AutoSave) {
            try {
                $Document.Save($Show.IsPresent)
            } catch {
                if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                    throw
                } else {
                    Write-Warning "Save-OfficeWord - Couldn't save $($Document.FilePath) Error: $($_.Exception.Message)"
                }
            } finally {
                #$Document.Dispose()
            }
        }
    }
    $Document.Dispose()
}