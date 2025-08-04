Clear-Host
# Simple example creating and saving a presentation
$presentation = New-OfficePowerPoint -FilePath "$PSScriptRoot\Documents\BasicExample.pptx"
Save-OfficePowerPoint -Presentation $presentation -Show
