Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot 'Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'LoadExample.pptx'
$presentation = New-OfficePowerPoint -FilePath $path
Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 | Out-Null
Save-OfficePowerPoint -Presentation $presentation

$loaded = Get-OfficePowerPoint -FilePath $path
Write-Host "Loaded presentation with $($loaded.Slides.Count) slide(s)."
