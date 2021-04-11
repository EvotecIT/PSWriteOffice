Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$PowerPoint = New-OfficePowerPoint -FilePath $PSScriptRoot\Documents\PowerPoint.pptx

Save-OfficePowerPoint -PowerPoint $PowerPoint -Show