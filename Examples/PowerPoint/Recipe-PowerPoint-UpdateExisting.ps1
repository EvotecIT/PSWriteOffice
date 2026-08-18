param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\PowerPoint')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$path = Join-Path $OutputDirectory 'PowerPoint-Updated-Existing.pptx'
PptNew -Path $path {
    PptSlide { PptTitle -Title 'FY24 Results'; PptTextBox -Text 'FY24 revenue is ready.' -X 80 -Y 150 -Width 500 -Height 60; PptNotes -Text 'Explain the FY24 result.' }
    PptSlide { PptTitle -Title 'FY24 Priorities'; PptBullets -Bullets 'Retain customers','Improve margin' -X 80 -Y 150 -Width 500 -Height 140 }
}

$presentation = Get-OfficePowerPoint -FilePath $path
try {
    $changes = Update-OfficePowerPointText -Presentation $presentation -OldValue 'FY24' -NewValue 'FY25' -IncludeNotes
    Save-OfficePowerPoint -Presentation $presentation
} finally {
    $presentation.Dispose()
}

[pscustomobject]@{
    Path    = $path
    Changes = $changes
} | Format-List
