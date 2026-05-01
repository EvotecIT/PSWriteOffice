$modulePath = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
    $env:PSWRITEOFFICE_MODULE_MANIFEST
} else {
    (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1')
}
if (-not (Get-Module -Name PSWriteOffice)) { Import-Module $modulePath -ErrorAction Stop }
$documents = Join-Path $PSScriptRoot 'Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'LayoutPlaceholderAliases.pptx'

New-OfficePowerPoint -Path $path {
    PptSlide {
        PptTitle -Title 'Alias Demo'
        $placeholders = PptLayoutPlaceholders
        Write-Host "Layout placeholders: $($placeholders.Count)"
    }
}

Write-Host "Presentation saved to $path"
