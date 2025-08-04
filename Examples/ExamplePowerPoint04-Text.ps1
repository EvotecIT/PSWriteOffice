Clear-Host
# Import-Module .\PSWriteOffice.psd1 -Force

$presentation = New-OfficePowerPoint -FilePath "$PSScriptRoot\Documents\ExamplePowerPoint4.pptx"

$slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Quarterly Report'
Add-OfficePowerPointTextBox -Slide $slide -Text 'Generated with PSWriteOffice'

Save-OfficePowerPoint -Presentation $presentation -Show
