Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force

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
