Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'PowerPoint-Charts.pptx'
$rows = @(
    [PSCustomObject]@{ Month = 'Jan'; MonthNumber = 1; Sales = 10; Profit = 4 }
    [PSCustomObject]@{ Month = 'Feb'; MonthNumber = 2; Sales = 14; Profit = 6 }
    [PSCustomObject]@{ Month = 'Mar'; MonthNumber = 3; Sales = 18; Profit = 8 }
)

$ppt = New-OfficePowerPoint -FilePath $path

$columnSlide = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
Set-OfficePowerPointSlideTitle -Slide $columnSlide -Title 'Column Chart' | Out-Null
Add-OfficePowerPointChart -Slide $columnSlide -Data $rows -CategoryProperty Month -SeriesProperty Sales, Profit -Title 'Sales vs Profit' | Out-Null

$pieSlide = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
Set-OfficePowerPointSlideTitle -Slide $pieSlide -Title 'Pie Chart' | Out-Null
Add-OfficePowerPointChart -Slide $pieSlide -Type Pie -Data $rows -CategoryProperty Month -SeriesProperty Sales -Title 'Sales Mix' | Out-Null

$scatterSlide = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
Set-OfficePowerPointSlideTitle -Slide $scatterSlide -Title 'Scatter Chart' | Out-Null
Add-OfficePowerPointChart -Slide $scatterSlide -Type Scatter -Data $rows -XProperty MonthNumber -YProperty Sales, Profit -Title 'Trend Scatter' | Out-Null

Save-OfficePowerPoint -Presentation $ppt
$ppt.Dispose()

Write-Host "Presentation saved to $path"
