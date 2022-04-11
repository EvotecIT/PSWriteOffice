function Remove-OfficeWordFooter {
    [cmdletBinding()]
    param(
        [parameter(Mandatory)][OfficeIMO.Word.WordDocument] $Document
    )
    try {
        [OfficeIMO.Word.WordFooter]::RemoveFooters($Document)
    } catch {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning -Message "Remove-OfficeWordFooter - Couldn't remove footer. Error: $($_.Exception.Message)"
        }
    }
}