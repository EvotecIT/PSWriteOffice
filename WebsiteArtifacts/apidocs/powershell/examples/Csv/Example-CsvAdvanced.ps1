Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Csv-Advanced.csv'
$rows = @(
    [PSCustomObject]@{ Name = 'Alpha'; Score = 92; Active = $true }
    [PSCustomObject]@{ Name = 'Beta'; Score = 76; Active = $true }
    [PSCustomObject]@{ Name = 'Gamma'; Score = 64; Active = $false }
)

$rows | ConvertTo-OfficeCsv -OutputPath $path -Delimiter ';' | Out-Null

Write-Host "CSV saved to $path"
Get-OfficeCsvData -Path $path -Delimiter ';' -AsHashtable | Format-Table