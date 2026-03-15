Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'PowerPoint-ThemeAndLayout.pptx'
$ppt = New-OfficePowerPoint -FilePath $path

$slide = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Theme Demo' | Out-Null

$layouts = Get-OfficePowerPointLayout -Presentation $ppt
$targetLayout = $layouts | Where-Object LayoutIndex -ne $slide.LayoutIndex | Select-Object -First 1

Set-OfficePowerPointThemeColor -Presentation $ppt -Colors @{
    Accent1 = '#C00000'
    Accent2 = '#00B0F0'
} -AllMasters
Set-OfficePowerPointThemeFonts -Presentation $ppt -MajorLatin 'Aptos' -MinorLatin 'Calibri' -AllMasters
Set-OfficePowerPointThemeName -Presentation $ppt -Name 'Contoso Theme' -AllMasters

if ($targetLayout.Type) {
    $slide | Set-OfficePowerPointSlideLayout -LayoutType $targetLayout.Type -Master $targetLayout.MasterIndex | Out-Null
} elseif ($targetLayout.Name) {
    $slide | Set-OfficePowerPointSlideLayout -LayoutName $targetLayout.Name -Master $targetLayout.MasterIndex | Out-Null
} else {
    $slide | Set-OfficePowerPointSlideLayout -Layout $targetLayout.LayoutIndex -Master $targetLayout.MasterIndex | Out-Null
}

$theme = Get-OfficePowerPointTheme -Presentation $ppt
$theme | Format-List

Save-OfficePowerPoint -Presentation $ppt
$ppt.Dispose()

Write-Host "Presentation saved to $path"
