function Remove-OfficeWordHeader {
    [cmdletBinding()]
    param(
        [DocumentFormat.OpenXml.Packaging.WordprocessingDocument] $Document
    )
    try {
        [OfficeImo.Headers]::RemoveHeaders($Document)
    } catch {
        Write-Warning -Message "Remove-OfficeWordHeader - Couldn't remove footer. Error: $($_.Exception.Message)"
    }
}