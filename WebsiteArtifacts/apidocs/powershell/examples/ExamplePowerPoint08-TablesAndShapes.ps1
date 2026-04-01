Import-Module (Join-Path $PSScriptRoot '..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot 'Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'ExamplePowerPoint8-TablesAndShapes.pptx'
$presentation = New-OfficePowerPoint -FilePath $path

$slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Tables & Shapes' | Out-Null

$data = @(
    [PSCustomObject]@{ Product = 'Alpha'; Qty = 12; Revenue = 1200 }
    [PSCustomObject]@{ Product = 'Beta'; Qty = 7; Revenue = 940 }
    [PSCustomObject]@{ Product = 'Gamma'; Qty = 20; Revenue = 1840 }
)

Add-OfficePowerPointTable -Slide $slide -Data $data -X 60 -Y 140 -Width 420 -Height 200 | Out-Null
Add-OfficePowerPointShape -Slide $slide -ShapeType Ellipse -X 520 -Y 140 -Width 140 -Height 140 -FillColor '#FFE699' -OutlineColor '#C65911' -OutlineWidth 1 | Out-Null
Add-OfficePowerPointTextBox -Slide $slide -Text 'Highlights' -X 530 -Y 300 -Width 120 -Height 40 | Out-Null

Save-OfficePowerPoint -Presentation $presentation
Write-Host "Presentation saved to $path"
