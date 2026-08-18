param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\PowerPoint')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$sourcePath = Join-Path $OutputDirectory 'PowerPoint-Reusable-Slides.pptx'
$targetPath = Join-Path $OutputDirectory 'PowerPoint-Combined-Deck.pptx'
PptNew -Path $sourcePath { PptSlide { PptTitle -Title 'Reusable Architecture'; PptTextBox -Text 'Shared platform diagram' -X 80 -Y 150 -Width 500 -Height 80 } }
PptNew -Path $targetPath { PptSlide { PptTitle -Title 'Customer Briefing'; PptTextBox -Text 'Prepared for review' -X 80 -Y 150 -Width 500 -Height 80 } }

$target = Get-OfficePowerPoint -FilePath $targetPath
try {
    Import-OfficePowerPointSlide -Presentation $target -SourcePath $sourcePath -SourceIndex 0 -InsertAt 1 | Out-Null
    Copy-OfficePowerPointSlide -Presentation $target -Index 0 -InsertAt 2 | Out-Null
    Add-OfficePowerPointSection -Presentation $target -Name 'Shared material' -StartSlideIndex 1 | Out-Null
    Save-OfficePowerPoint -Presentation $target

    [pscustomobject]@{
        Path     = $targetPath
        Slides   = $target.Slides.Count
        Sections = @(Get-OfficePowerPointSection -Presentation $target).Count
    } | Format-List
} finally {
    $target.Dispose()
}
