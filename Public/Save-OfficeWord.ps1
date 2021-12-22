function Save-OfficeWord {
    [cmdletBinding()]
    param(
        [alias('WordDocument')][DocumentFormat.OpenXml.Packaging.WordprocessingDocument] $Document,
        [switch] $Show,
        [string] $FilePath,
        [int] $Retry = 2
    )
    if (-not $Document) {
        Write-Warning "Save-OfficeWord - Couldn't save Word Document. Document is null."
        return
    }
    $Saved = $false
    $Count = 0
    if (-not $Document.FilePath -and -not $FilePath) {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning "Save-OfficeWord - Couldn't save Word Document. No file path provided."
            return
        }
    }
    while ($Count -le $Retry -and $Saved -eq $false) {
        $Count++
        if ($FilePath) {
            # File path was given so we use SaveAs
            try {
                $NewDocument = $Document.SaveAs($FilePath)
                $Saved = $true
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
                    $Saved = $true
                } catch {
                    if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                        throw
                    } else {
                        Write-Warning "Save-OfficeWord - Couldn't save $FilePath. Error: $($_.Exception.Message)"
                    }
                }
            }
        }
        if (-not $Saved) {
            if ($Retry -ge $Count) {
                $FilePath = [io.path]::GetTempFileName().Replace('.tmp', '.docx')
                Write-Warning -Message "Save-OfficeWord - Couldn't save using provided file name, retrying with $FilePath"
            } else {
                Write-Warning -Message "Save-OfficeWord - Couldn't save using provided file name. Run out of retries ($Count / $Retry)."
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