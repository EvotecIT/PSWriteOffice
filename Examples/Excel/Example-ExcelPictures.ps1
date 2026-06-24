$modulePath = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
    $env:PSWRITEOFFICE_MODULE_MANIFEST
} else {
    (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1')
}
if (-not (Get-Module -Name PSWriteOffice)) { Import-Module $modulePath -ErrorAction Stop }

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Excel-Pictures.xlsx'
$officeimoRoot = Join-Path $PSScriptRoot '..\..\..\OfficeIMO'
$imageCandidates = @(
    (Join-Path (Join-Path $officeimoRoot 'Assets') 'OfficeIMO.png')
    (Join-Path (Join-Path $officeimoRoot 'OfficeIMO.Tests\Images') 'EvotecLogo.png')
)
$imagePath = $imageCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $imagePath) {
    throw 'Could not find a sample image. Run this example from an EvotecIT checkout with OfficeIMO next to PSWriteOffice.'
}

New-OfficeExcel -Path $path {
    ExcelSheet 'Pictures' {
        ExcelCell -Address 'A1' -Value 'Range anchored image'
        ExcelCell -Address 'E1' -Value 'Scaled image'
        ExcelCell -Address 'E8' -Value 'Rotated image'

        ExcelImage -Path $imagePath -Range 'A2:C12' -Name 'HeaderLogo' -AltText 'Company logo pinned to A2 through C12' -Placement MoveAndSize
        ExcelImage -Path $imagePath -Address 'E2' -ScalePercent 20 -Name 'ScaledLogo' -AltText 'Logo scaled to 20 percent'
        ExcelImage -Path $imagePath -Address 'E9' -Width 120 -Height 48 -RotationDegrees 12 -Name 'RotatedLogo' -AltText 'Logo rotated by 12 degrees'

        ExcelColumn -ColumnName A -Width 18
        ExcelColumn -ColumnName B -Width 18
        ExcelColumn -ColumnName C -Width 18
        ExcelColumn -ColumnName E -Width 24
    }
} | Out-Null

Write-Host "Workbook saved to $path"
