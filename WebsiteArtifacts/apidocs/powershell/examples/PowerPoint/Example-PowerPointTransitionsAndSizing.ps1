Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'PowerPoint-TransitionsAndSizing.pptx'

$presentation = New-OfficePowerPoint -FilePath $path
Set-OfficePowerPointSlideSize -Presentation $presentation -Preset Screen16x9 | Out-Null

$intro = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Set-OfficePowerPointSlideTitle -Slide $intro -Title 'Executive Summary' | Out-Null
Set-OfficePowerPointSlideTransition -Slide $intro -Transition Fade | Out-Null

$details = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Set-OfficePowerPointSlideTitle -Slide $details -Title 'Details' | Out-Null
Set-OfficePowerPointSlideTransition -Slide $details -Transition Morph | Out-Null

Save-OfficePowerPoint -Presentation $presentation

Write-Host "Presentation saved to $path"
