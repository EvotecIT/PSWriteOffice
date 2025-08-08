Clear-Host
# Import-Module .\\PSWriteOffice.psd1 -Force

$presentation = New-OfficePowerPoint -FilePath "$PSScriptRoot\Documents\ExamplePowerPoint6.pptx"
Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 | Out-Null
Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 | Out-Null
Remove-OfficePowerPointSlide -Presentation $presentation -Index 0
Save-OfficePowerPoint -Presentation $presentation -Show
