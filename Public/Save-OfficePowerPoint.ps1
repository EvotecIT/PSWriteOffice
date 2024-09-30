function Save-OfficePowerPoint {
    [cmdletBinding()]
    param(
        [Parameter(Mandatory)][ShapeCrawler.Presentation] $Presentation,
        [switch] $Show
    )

    if (-not $Script:PowerPointConfiguration) {
        Write-Warning -Message "Save-OfficePowerPoint - Couldn't save PowerPoint Presentation. Presentation is null."
        return
    }
    if ($Script:PowerPointConfiguration.State -eq 'New') {
        $Presentation.SaveAs($Script:PowerPointConfiguration.FilePath)
    } else {
        $Presentation.Save()
    }
    if ($Show) {
        Invoke-Item -Path $Script:PowerPointConfiguration.FilePath
    }
}