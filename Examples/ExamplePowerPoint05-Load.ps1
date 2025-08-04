Clear-Host
# Create and load a PowerPoint presentation
$path = "$PSScriptRoot\Documents\LoadExample.pptx"
$presentation = New-OfficePowerPoint -FilePath $path
Save-OfficePowerPoint -Presentation $presentation
$loaded = Get-OfficePowerPoint -FilePath $path
$loaded
