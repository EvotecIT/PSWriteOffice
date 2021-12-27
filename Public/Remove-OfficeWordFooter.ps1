function Remove-OfficeWordFooter {
    [cmdletBinding()]
    param(
        [DocumentFormat.OpenXml.Packaging.WordprocessingDocument] $Document
    )
    try {
        [OfficeIMO.Word.Footers]::RemoveFooters($Document)
    } catch {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning -Message "Remove-OfficeWordFooter - Couldn't remove footer. Error: $($_.Exception.Message)"
        }
    }
}