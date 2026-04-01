Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot 'Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'BasicExample.pptx'
$presentation = New-OfficePowerPoint -FilePath $path

Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 | Out-Null
Save-OfficePowerPoint -Presentation $presentation

Write-Host "Presentation saved to $path"
