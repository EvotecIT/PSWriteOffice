Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'PowerPoint-CopySlides.pptx'
$presentation = New-OfficePowerPoint -FilePath $path

$intro = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Set-OfficePowerPointSlideTitle -Slide $intro -Title 'Executive Summary'
Add-OfficePowerPointTextBox -Slide $intro -Text 'Quarterly revenue and margin summary' -X 80 -Y 150 -Width 360 -Height 60 | Out-Null
Set-OfficePowerPointNotes -Slide $intro -Text 'Use this for board prep.'

$closing = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Set-OfficePowerPointSlideTitle -Slide $closing -Title 'Appendix'

Copy-OfficePowerPointSlide -Presentation $presentation -Index 0 -InsertAt 1 | Out-Null

Save-OfficePowerPoint -Presentation $presentation

Write-Host "Presentation saved to $path"
