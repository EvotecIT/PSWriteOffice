Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'PowerPoint-BackgroundsAndLayout.pptx'
$imagePath = Join-Path $documents 'PowerPoint-Background.bmp'

[byte[]] $bytes = 0x42, 0x4D, 0x3A, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x36, 0x00, 0x00, 0x00, 0x28, 0x00,
    0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00,
    0x00, 0x00, 0x01, 0x00, 0x18, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x13, 0x0B,
    0x00, 0x00, 0x13, 0x0B, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF,
    0xFF, 0x00
[System.IO.File]::WriteAllBytes($imagePath, $bytes)

$ppt = New-OfficePowerPoint -FilePath $path
Set-OfficePowerPointSlideSize -Presentation $ppt -WidthCm 30 -HeightCm 20 | Out-Null

$slide1 = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
Set-OfficePowerPointSlideTitle -Slide $slide1 -Title 'Content Grid' | Out-Null
Set-OfficePowerPointBackground -Slide $slide1 -Color '#F4F7FB' | Out-Null

$columns = @(Get-OfficePowerPointLayoutBox -Presentation $ppt -ColumnCount 2 -MarginCm 1.5 -GutterCm 1.0)
foreach ($index in 0..($columns.Count - 1)) {
    $box = $columns[$index]
    Add-OfficePowerPointTextBox -Slide $slide1 -Text "Column $($index + 1)" -X $box.LeftPoints -Y $box.TopPoints -Width $box.WidthPoints -Height 48 | Out-Null
}

$slide2 = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
Set-OfficePowerPointSlideTitle -Slide $slide2 -Title 'Image Background' | Out-Null
Set-OfficePowerPointBackground -Slide $slide2 -ImagePath $imagePath | Out-Null

Save-OfficePowerPoint -Presentation $ppt
$ppt.Dispose()

Write-Host "Presentation saved to $path"
