function Save-OfficePowerPoint {
    [cmdletBinding()]
    param(
        $PowerPoint,
        [switch] $Show
    )
    $FilePath = $PowerPoint.FilePath
    $PowerPoint.Close()
    if ($Show) {
        Invoke-Item -LiteralPath $FilePath
    }
}