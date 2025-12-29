Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot 'Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'ExamplePowerPoint4.pptx'
$presentation = New-OfficePowerPoint -FilePath $path

$slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Quarterly Report' | Out-Null
Add-OfficePowerPointTextBox -Slide $slide -Text 'Generated with PSWriteOffice' -X 90 -Y 160 -Width 320 -Height 50 | Out-Null

Save-OfficePowerPoint -Presentation $presentation
Write-Host "Presentation saved to $path"
