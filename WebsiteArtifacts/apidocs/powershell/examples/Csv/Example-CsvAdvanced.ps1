$modulePath = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
    $env:PSWRITEOFFICE_MODULE_MANIFEST
} else {
    (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1')
}
if (-not (Get-Module -Name PSWriteOffice)) { Import-Module $modulePath -ErrorAction Stop }
$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Csv-Advanced.csv'
$rows = @(
    [PSCustomObject]@{ Name = 'Alpha'; Score = 92; Active = $true }
    [PSCustomObject]@{ Name = 'Beta'; Score = 76; Active = $true }
    [PSCustomObject]@{ Name = 'Gamma'; Score = 64; Active = $false }
)

$rows | Export-OfficeCsv -Path $path -Delimiter ';'

Write-Host "CSV saved to $path"
Import-OfficeCsv -Path $path -Delimiter ';' -AsHashtable | Format-Table
