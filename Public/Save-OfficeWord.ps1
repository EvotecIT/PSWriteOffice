function Save-OfficeWord {
    [cmdletBinding()]
    param(
        [alias('WordDocument')][DocumentFormat.OpenXml.Packaging.WordprocessingDocument] $Document,
        [switch] $Show,
        [string] $FilePath
    )
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
            $NewDocument = $Document.SaveAs($FilePath)
        } catch {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw
            } else {
                Write-Warning "Save-OfficeWord - Couldn't save $FilePath. Error: $($_.Exception.Message)"
            }
        }
    } else {
        $FilePath = $Document.FilePath
        if (-not $Document.AutoSave) {
            try {
                $Document.Save()
            } catch {
                if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                    throw
                } else {
                    Write-Warning "Save-OfficeWord - Couldn't save $FilePath. Error: $($_.Exception.Message)"
                }
            }
        }
    }
    try {
        $Document.Close()
    } catch {
        if ( $_.Exception.InnerException.Message -eq "Memory stream is not expandable.") {
            # we swallow this exception because it only fails on PS 7.
        } else {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw
            } else {
                Write-Warning "Save-OfficeWord - Couldn't close document for $FilePath. Error: $($_.Exception.Message)"
            }
        }
    }

    if ($NewDocument) {
        try {
            $NewDocument.Close()
        } catch {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw
            } else {
                Write-Warning "Save-OfficeWord - Couldn't close document for $FilePath. Error: $($_.Exception.Message)"
            }
        }
    }
    if ($Show) {
        try {
            Invoke-Item -LiteralPath $FilePath -ErrorAction Stop
        } catch {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw
            } else {
                Write-Warning "Save-OfficeWord - Couldn't open $FilePath. Error: $($_.Exception.Message)"
            }
        }
    }
}