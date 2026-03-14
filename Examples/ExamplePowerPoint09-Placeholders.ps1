Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot 'Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'ExamplePowerPoint9-Placeholders.pptx'
$presentation = New-OfficePowerPoint -FilePath $path

$layouts = Get-OfficePowerPointLayout -Presentation $presentation
$layout = $layouts | Where-Object { $_.Type } | Select-Object -First 1
if (-not $layout) {
    $layout = $layouts | Select-Object -First 1
}

$slide = if ($layout.Type) {
    Add-OfficePowerPointSlide -Presentation $presentation -LayoutType $layout.Type -Master $layout.MasterIndex
} elseif ($layout.Name) {
    Add-OfficePowerPointSlide -Presentation $presentation -LayoutName $layout.Name -Master $layout.MasterIndex
} else {
    Add-OfficePowerPointSlide -Presentation $presentation -Layout $layout.LayoutIndex -Master $layout.MasterIndex
}

Set-OfficePowerPointPlaceholderText -Slide $slide -PlaceholderType Title -Text 'Status Update' | Out-Null

$layoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder -Slide $slide
$placeholder = $layoutPlaceholders | Where-Object { $_.PlaceholderType } | Select-Object -First 1
if ($placeholder) {
    $placeholderType = $placeholder.PlaceholderType.Value
    Set-OfficePowerPointLayoutPlaceholderBounds -Presentation $presentation -Master $layout.MasterIndex -Layout $layout.LayoutIndex `
        -PlaceholderType $placeholderType -Index $placeholder.PlaceholderIndex -Left 60 -Top 140 -Width 520 -Height 240 | Out-Null
    Set-OfficePowerPointLayoutPlaceholderTextMargins -Presentation $presentation -Master $layout.MasterIndex -Layout $layout.LayoutIndex `
        -PlaceholderType $placeholderType -Index $placeholder.PlaceholderIndex -Left 12 -Top 8 -Right 12 -Bottom 8 | Out-Null
    Set-OfficePowerPointLayoutPlaceholderTextStyle -Presentation $presentation -Master $layout.MasterIndex -Layout $layout.LayoutIndex `
        -PlaceholderType $placeholderType -Index $placeholder.PlaceholderIndex -Style Body -FontSize 18 -Bold $true | Out-Null
}

Save-OfficePowerPoint -Presentation $presentation
Write-Host "Presentation saved to $path"
