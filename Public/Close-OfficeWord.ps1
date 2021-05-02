function Close-OfficeWord {
    [cmdletBinding()]
    param(
        [alias('WordDocument')] $Document
    )
    try {
        $Document.Close()
    } catch {
        if ( $_.Exception.InnerException.Message -eq "Memory stream is not expandable.") {
            # we swallow this exception because it only fails on PS 7.
        } else {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw
            } else {
                Write-Warning "Close-OfficeWord - Couldn't close document properly. Error: $($_.Exception.Message)"
            }
        }
    }
}