function Save-OfficeWordDocument {
    [cmdletBinding()]
    param(
        [alias('WordDocument')] $Document,
        [switch] $Show
    )
    $FilePath = $Document.FilePath
    $Document.Close()
    if ($Show) {
        Invoke-Item -LiteralPath $FilePath
    }
}