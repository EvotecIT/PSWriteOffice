function New-OfficePowerPoint {
    [cmdletBinding()]
    param(
        [string] $FilePath
    )

    $Script:PowerPointConfiguration = @{
        State    = 'New'
        FilePath = $FilePath
    }
    $Presentation = [ShapeCrawler.Presentation]::New()
    $Presentation
}