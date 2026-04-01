Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot 'Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'ExamplePowerPoint1.pptx'
$presentation = New-OfficePowerPoint -FilePath $path

$slide1 = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Set-OfficePowerPointSlideTitle -Slide $slide1 -Title 'Status Update' | Out-Null
Add-OfficePowerPointTextBox -Slide $slide1 -Text 'Generated with PSWriteOffice' -X 80 -Y 150 -Width 320 -Height 40 | Out-Null
Add-OfficePowerPointShape -Slide $slide1 -ShapeType Rectangle -X 80 -Y 210 -Width 320 -Height 120 -FillColor '#DDEEFF' -OutlineColor '#4472C4' -OutlineWidth 1 | Out-Null

$slide2 = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Set-OfficePowerPointSlideTitle -Slide $slide2 -Title 'Next Steps' | Out-Null
Add-OfficePowerPointTextBox -Slide $slide2 -Text '1. Review numbers  2. Plan Q1  3. Ship' -X 80 -Y 150 -Width 360 -Height 80 | Out-Null

Save-OfficePowerPoint -Presentation $presentation
Write-Host "Presentation saved to $path"
