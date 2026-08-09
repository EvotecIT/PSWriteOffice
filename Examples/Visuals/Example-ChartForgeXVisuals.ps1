$ErrorActionPreference = 'Stop'

Import-Module ImagePlayground -ErrorAction Stop
Import-Module PSWriteOffice -ErrorAction Stop

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$chart = New-ImageTopology -Node @(
    New-ImageTopologyNode -Id api -Label API -Detail (New-ImageTopologyNodeDetail -Label Runtime -Value '.NET 10')
    New-ImageTopologyNode -Id db -Label Database
) -Edge (
    New-ImageTopologyEdge -SourceNodeId api -TargetNodeId db -Label SQL -PreferredLength 180 -TargetMarker Arrow
) -LayoutPreset Presentation -FilePath (Join-Path $documents 'service-map.svg') -PassThru

$artifact = $chart |
    ConvertTo-ImageVisualArtifact -Id service-map -AccessibleDescription 'API service connects to the database.'
$svgPath = Join-Path $documents 'service-map-office.svg'
$artifact | Export-ImageVisualArtifact -FilePath $svgPath
$officeVisual = $svgPath | ConvertTo-OfficeVisual -Width 420 -SvgPolicy RasterizeWhenNeeded `
    -Id service-map -Title 'Service Map' -AlternativeText 'API service connects to the database.'

New-OfficeWord -Path (Join-Path $documents 'service-map.docx') {
    WordSection { WordParagraph { $officeVisual | Add-OfficeWordVisual | Out-Null } }
} | Out-Null

New-OfficeExcel -Path (Join-Path $documents 'service-map.xlsx') {
    Add-OfficeExcelSheet -Name Dashboard -Content {
        $officeVisual | Add-OfficeExcelVisual -Address B2 | Out-Null
    }
} | Out-Null

New-OfficePowerPoint -Path (Join-Path $documents 'service-map.pptx') {
    PptSlide { $officeVisual | Add-OfficePowerPointVisual -X 48 -Y 72 | Out-Null }
} | Out-Null

New-OfficePdf -Path (Join-Path $documents 'service-map.pdf') {
    $officeVisual | Add-OfficePdfVisual -Align Center
} | Out-Null

$officeVisual.Report
