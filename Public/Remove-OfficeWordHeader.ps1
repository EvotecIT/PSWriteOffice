function Remove-OfficeWordHeader {
    [cmdletBinding()]
    param(
        [DocumentFormat.OpenXml.Packaging.WordprocessingDocument] $Document
    )
    try {
        [OfficeImo.Headers]::RemoveHeaders($Document)
    } catch {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning -Message "Remove-OfficeWordHeader - Couldn't remove footer. Error: $($_.Exception.Message)"
        }
    }
}