function Remove-OfficeWordHeader {
    [cmdletBinding()]
    param(
        [parameter(Mandatory)][OfficeIMO.Word.WordDocument] $Document
    )
    try {
        if ($Document) {
            [OfficeIMO.Word.WordHeader]::RemoveHeaders($Document)
        } else {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw "Couldn't remove footer. Document not provided."
            } else {
                Write-Warning -Message "Remove-OfficeWordHeader - Couldn't remove footer. Document not provided."
            }
        }
    } catch {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning -Message "Remove-OfficeWordHeader - Couldn't remove footer. Error: $($_.Exception.Message)"
        }
    }
}