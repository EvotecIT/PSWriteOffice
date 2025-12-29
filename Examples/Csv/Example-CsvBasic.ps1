Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Data.csv'
$rows = @(
    [PSCustomObject]@{ Name = 'Alpha'; Value = 1 }
    [PSCustomObject]@{ Name = 'Beta'; Value = 2 }
)

$rows | ConvertTo-OfficeCsv -OutputPath $path | Out-Null

$data = Get-OfficeCsvData -Path $path
$data | Format-Table
