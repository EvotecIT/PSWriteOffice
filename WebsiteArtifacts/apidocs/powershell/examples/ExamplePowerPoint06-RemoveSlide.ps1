$modulePath = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
    $env:PSWRITEOFFICE_MODULE_MANIFEST
} else {
    (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1')
}
if (-not (Get-Module -Name PSWriteOffice)) { Import-Module $modulePath -ErrorAction Stop }
$documents = Join-Path $PSScriptRoot 'Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'ExamplePowerPoint6.pptx'
$presentation = New-OfficePowerPoint -FilePath $path
Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 | Out-Null
Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 | Out-Null

Remove-OfficePowerPointSlide -Presentation $presentation -Index 0
Save-OfficePowerPoint -Presentation $presentation

Write-Host "Presentation saved to $path"
