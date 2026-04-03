Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$path = Join-Path $PSScriptRoot 'Example-WordTableCells.docx'
$imagePath = Join-Path $PSScriptRoot 'Example-WordTableCells.png'
$rows = @(
    [pscustomobject]@{
        Topic   = 'Release readiness'
        Details = 'Pending'
    }
)

$nestedRows = @(
    [pscustomobject]@{ Step = 'Validate'; State = 'Ready' }
    [pscustomobject]@{ Step = 'Release'; State = 'Queued' }
)

Add-Type -AssemblyName System.Drawing
$bitmap = [System.Drawing.Bitmap]::new(48, 48)
try {
    for ($x = 0; $x -lt $bitmap.Width; $x++) {
        for ($y = 0; $y -lt $bitmap.Height; $y++) {
            $bitmap.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(255, 33, 150, 243))
        }
    }

    $bitmap.Save($imagePath, [System.Drawing.Imaging.ImageFormat]::Png)
} finally {
    $bitmap.Dispose()
}

New-OfficeWord -Path $path {
    WordTable -Data $rows -Style GridTable1LightAccent1 {
        WordTableCell -Row 1 -Column 0 {
            WordParagraph {
                WordText 'Checklist'
            }

            WordImage -Path $imagePath -Width 36 -Height 36 -Description 'Status icon'

            WordList {
                WordListItem 'Confirm issue coverage'
                WordListItem 'Stage release notes'
            }
        }

        WordTableCell -Row 1 -Column 1 {
            WordTable -Data $nestedRows -Style TableGrid
        }
    }
} | Out-Null

Write-Host "Document saved to $path"
Write-Host "Image fixture saved to $imagePath"
