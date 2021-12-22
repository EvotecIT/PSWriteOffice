function Remove-OfficeWordFooter {
    [cmdletBinding()]
    param(
        [DocumentFormat.OpenXml.Packaging.WordprocessingDocument] $Document
    )
    try {
        [OfficeImo.Footers]::RemoveFooters($Document)
    } catch {
        Write-Warning -Message "Remove-OfficeWordFooter - Couldn't remove footer. Error: $($_.Exception.Message)"
    }
}