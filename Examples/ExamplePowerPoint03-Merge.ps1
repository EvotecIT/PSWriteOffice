Clear-Host
#Import-Module .\PSWriteOffice.psd1 -Force

$targetPath = "$PSScriptRoot\Documents\MergedPresentation.pptx"
$sourcePath = "$PSScriptRoot\Documents\SourcePresentation.pptx"

$Target = New-OfficePowerPoint -FilePath $targetPath
$Source = New-OfficePowerPoint -FilePath $sourcePath
Add-OfficePowerPointSlide -Presentation $Source -Layout 1
Save-OfficePowerPoint -Presentation $Source

Merge-OfficePowerPoint -Presentation $Target -FilePath $sourcePath
Save-OfficePowerPoint -Presentation $Target -Show

