Clear-Host
# Demonstrates using WhatIf when saving a presentation
$path = "$PSScriptRoot\Documents\WhatIfExample.pptx"
$presentation = New-OfficePowerPoint -FilePath $path
Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 | Out-Null
Save-OfficePowerPoint -Presentation $presentation -WhatIf
