Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'PowerPoint-ModifyExistingShapes.pptx'

$initialRows = @(
    [PSCustomObject]@{ Metric = 'Risk'; State = 'Open' }
    [PSCustomObject]@{ Metric = 'Quality'; State = 'Watching' }
)

$presentation = New-OfficePowerPoint -FilePath $path
try {
    $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
    Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Release readiness' | Out-Null
    Add-OfficePowerPointTextBox -Slide $slide -Text 'Status marker: Draft release' -X 70 -Y 110 -Width 420 -Height 45 | Out-Null
    Add-OfficePowerPointTable -Slide $slide -InputObject $initialRows -X 70 -Y 180 -Width 500 -Height 170 | Out-Null
} finally {
    Close-OfficePowerPoint -Presentation $presentation -Save
}

# Second pass: find existing shapes, then modify their content directly.
$deck = Get-OfficePowerPoint -FilePath $path
try {
    Find-OfficePowerPointShape -Presentation $deck -Text 'Status marker' -Kind TextBox |
        Set-OfficePowerPointShapeText -Text 'Status marker: Ready for launch' |
        Out-Null

    $readinessTable = Find-OfficePowerPointShape -Presentation $deck -Text 'Risk' -Kind Table | Select-Object -First 1

    $readinessTable |
        Add-OfficePowerPointTableRow -Values 'Latency', 'Investigating' |
        Out-Null

    $readinessTable |
        Add-OfficePowerPointTableRow -Values ([ordered]@{
            Metric = 'Documentation'
            State  = 'Ready'
        }) |
        Out-Null

    $readinessTable |
        Set-OfficePowerPointTableCell -Row 1 -Column 1 -Text 'Mitigating' |
        Out-Null
} finally {
    Close-OfficePowerPoint -Presentation $deck -Save
}

Write-Host "Updated PowerPoint deck saved to $path"
Write-Host 'Matching shapes:'
$reloaded = Get-OfficePowerPoint -FilePath $path
try {
    Find-OfficePowerPointShape -Presentation $reloaded -Text 'Ready' |
        Select-Object SlideIndex, ShapeIndex, Kind, Text |
        Format-Table
} finally {
    Close-OfficePowerPoint -Presentation $reloaded
}
