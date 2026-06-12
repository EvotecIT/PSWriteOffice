BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Visio cmdlets' {
    It 'creates, loads, and inspects a Visio document' {
        $path = Join-Path $TestDrive 'diagram.vsdx'

        $document = New-OfficeVisio -Path $path -Title 'Visio smoke' -Author 'PSWriteOffice' -PassThru
        $document.Pages[0].AddRectangle(2, 2, 2, 1, 'Visio smoke') | Out-Null
        $document | Save-OfficeVisio -Path $path | Out-Null

        Test-Path $path | Should -BeTrue

        $loaded = Get-OfficeVisio -Path $path
        $loaded.Pages.Count | Should -Be 1

        $info = Get-OfficeVisioInfo -Path $path
        $info.Title | Should -Be 'Visio smoke'
        $info.ShapeCount | Should -Be 1

        $text = Get-OfficeVisioInfo -Path $path -AsText
        $text | Should -Match 'document.shapeCount=1'
        $text | Should -Match 'Visio smoke'
    }

    It 'exports Visio documents to SVG and PNG' {
        $path = Join-Path $TestDrive 'export.vsdx'
        $svgPath = Join-Path $TestDrive 'export.svg'
        $pngPath = Join-Path $TestDrive 'export.png'

        $document = New-OfficeVisio -Path $path -PassThru
        $document.Pages[0].AddRectangle(2, 2, 2, 1, 'SVG smoke') | Out-Null
        $document | Save-OfficeVisio -Path $path | Out-Null

        ConvertTo-OfficeVisioSvg -Path $path -OutputPath $svgPath |
            Should -BeOfType System.IO.FileInfo
        $svg = Get-Content -Path $svgPath -Raw
        $svg | Should -Match '<svg'
        $svg | Should -Match 'SVG smoke'

        ConvertTo-OfficeVisioPng -Path $path -OutputPath $pngPath -NoText |
            Should -BeOfType System.IO.FileInfo
        $bytes = [System.IO.File]::ReadAllBytes($pngPath)
        $bytes.Length | Should -BeGreaterThan 8
        $bytes[0] | Should -Be 137
        $bytes[1] | Should -Be 80
        $bytes[2] | Should -Be 78
        $bytes[3] | Should -Be 71
    }

    It 'creates a diagram through the Visio DSL' {
        $path = Join-Path $TestDrive 'dsl.vsdx'
        $svgPath = Join-Path $TestDrive 'dsl.svg'

        New-OfficeVisio -Path $path -Title 'DSL diagram' -RequestRecalcOnOpen {
            VisioRectangle -Key web -Text 'Web' -X 1.5 -Y 4 -Width 1.5 -Height 0.8 -FillColor LightBlue -LineColor SteelBlue
            VisioDiamond -Key decision -Text 'Ready?' -X 4 -Y 4 -Width 1.2 -Height 1 -FillColor '#FFF2CC' -LineColor '#B45309'
            VisioRectangle -Key api -Text 'API' -X 6.2 -Y 4 -Width 1.5 -Height 0.8 -FillColor LightGreen -LineColor SeaGreen
            VisioConnector -From web -To decision -Kind RightAngle -EndArrow Triangle -Label 'check'
            VisioConnector -From decision -To api -Kind RightAngle -EndArrow Triangle -Label 'ship'
            VisioPage 'Operations' {
                VisioTextBox 'Generated with PSWriteOffice Visio DSL' -X 2 -Y 2 -Width 4 -Height 0.5
                VisioEllipse -Key ops -Text 'Ops' -X 4 -Y 4 -Width 1.4 -Height 0.8 -FillColor WhiteSmoke -LineColor Gray
            }
        } | Out-Null

        $info = Get-OfficeVisioInfo -Path $path
        $info.Pages.Count | Should -Be 2
        $info.ShapeCount | Should -Be 5
        $info.ConnectorCount | Should -Be 2

        ConvertTo-OfficeVisioSvg -Path $path -OutputPath $svgPath | Out-Null
        $svg = Get-Content -Path $svgPath -Raw
        $svg | Should -Match 'Web'
        $svg | Should -Match 'Ready'
        $svg | Should -Match 'API'
    }

    It 'targets nested Visio page scopes and generates unique unkeyed stencil ids' {
        $path = Join-Path $TestDrive 'nested-stencils.vsdx'

        New-OfficeVisio -Path $path -Title 'Nested stencil diagram' {
            Import-OfficeVisioStencil -BuiltIn Flowchart -Name Flow -Default
            VisioPage 'Nested' {
                VisioStencil -Stencil process -Text 'First process' -X 2 -Y 3
                VisioStencil -Stencil process -Text 'Second process' -X 5 -Y 3
            }
        } | Out-Null

        $document = Get-OfficeVisio -Path $path
        @($document.Pages).Count | Should -Be 2
        @($document.Pages[0].Shapes).Count | Should -Be 0

        $nestedShapes = @($document.Pages[1].Shapes)
        $nestedShapes.Count | Should -Be 2
        @($nestedShapes | Select-Object -ExpandProperty Id -Unique).Count | Should -Be 2
        @($nestedShapes | Select-Object -ExpandProperty Text) | Should -Contain 'First process'
        @($nestedShapes | Select-Object -ExpandProperty Text) | Should -Contain 'Second process'
    }

    It 'keeps DSL connector key resolution scoped to the active page' {
        $path = Join-Path $TestDrive 'cross-page-connector.vsdx'

        {
            New-OfficeVisio -Path $path -Title 'Cross-page connector' {
                VisioRectangle -Key web -Text 'Web' -X 1.5 -Y 4 -Width 1.5 -Height 0.8
                VisioPage 'Operations' {
                    VisioEllipse -Key ops -Text 'Ops' -X 4 -Y 4 -Width 1.4 -Height 0.8
                    VisioConnector -From web -To ops
                }
            } | Out-Null
        } | Should -Throw "*Visio shape 'web' was not found on the active Visio page*"
    }

    It 'searches built-in stencil catalogs and creates a stencil-based DSL diagram' {
        $path = Join-Path $TestDrive 'stencil-dsl.vsdx'
        $svgPath = Join-Path $TestDrive 'stencil-dsl.svg'

        $flow = Get-OfficeVisioStencilCatalog -BuiltIn Flowchart
        $flow.Name | Should -Be 'Flowchart'
        (Find-OfficeVisioStencil -Catalog $flow -Query process -First 1).Id | Should -Be 'flow.process'

        New-OfficeVisio -Path $path -Title 'Stencil DSL diagram' -UseMastersByDefault -RequestRecalcOnOpen {
            Import-OfficeVisioStencil -BuiltIn Flowchart -Name Flow -Default
            VisioStencil -Catalog Flow -Stencil process -Key intake -Text 'Intake' -X 1.5 -Y 4 -FillColor '#E0F2FE' -LineColor '#0369A1'
            VisioStencil -Catalog Flow -Stencil decision -Key review -Text 'Review?' -X 4 -Y 4 -Width 1.6 -Height 1.1 -FillColor '#FEF3C7' -LineColor '#B45309'
            VisioStencil -Catalog Flow -Stencil data -Key archive -Text 'Archive' -X 6.5 -Y 4 -FillColor '#DCFCE7' -LineColor '#15803D'
            VisioConnector -From intake -To review -Kind RightAngle -EndArrow Triangle -Label 'submit'
            VisioConnector -From review -To archive -Kind RightAngle -EndArrow Triangle -Label 'store'
        } | Out-Null

        $info = Get-OfficeVisioInfo -Path $path
        $info.ShapeCount | Should -Be 3
        $info.ConnectorCount | Should -Be 2

        ConvertTo-OfficeVisioSvg -Path $path -OutputPath $svgPath | Out-Null
        $svg = Get-Content -Path $svgPath -Raw
        $svg | Should -Match 'Intake'
        $svg | Should -Match 'Review'
        $svg | Should -Match 'Archive'
    }

    It 'loads a package-backed stencil catalog and uses it in the DSL' {
        $repoRoot = if ($env:EVOTEC_GITHUB_ROOT) {
            $env:EVOTEC_GITHUB_ROOT
        } elseif ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
            'C:\Support\GitHub'
        } else {
            Join-Path $HOME 'Support/GitHub'
        }

        $templatePath = Join-Path $repoRoot 'OfficeIMO\Assets\VisioTemplates\DrawingWithShapes.vsdx'
        if (-not (Test-Path -LiteralPath $templatePath)) {
            Set-ItResult -Skipped -Because "OfficeIMO Visio template fixture was not found at $templatePath."
            return
        }

        $path = Join-Path $TestDrive 'package-stencil.vsdx'
        $catalog = Get-OfficeVisioStencilCatalog -Path $templatePath -CatalogName 'OfficeIMO Template' -IncludeUnsupportedMasters
        if ($catalog.Shapes.Count -eq 0) {
            Set-ItResult -Skipped -Because "OfficeIMO Visio template fixture did not expose package masters."
            return
        }

        $stencil = Find-OfficeVisioStencil -Catalog $catalog -First 1
        New-OfficeVisio -Path $path -Title 'Package stencil DSL' -UseMastersByDefault {
            Import-OfficeVisioStencil -Catalog $catalog -Name Template -Default
            VisioStencil -Stencil $stencil.Id -Key imported -Text 'Imported stencil' -X 2.5 -Y 3
        } | Out-Null

        $info = Get-OfficeVisioInfo -Path $path
        $info.ShapeCount | Should -Be 1
        $info.Title | Should -Be 'Package stencil DSL'
    }

    It 'preserves stencil default units when one dimension is omitted' {
        $path = Join-Path $TestDrive 'metric-stencil.vsdx'
        $visioAssembly = Get-TestPSWriteOfficeAssembly -Name 'OfficeIMO.Visio' -CommandName 'New-OfficeVisio'
        $stencilType = $visioAssembly.GetType('OfficeIMO.Visio.Stencils.VisioStencilShape', $true)
        $catalogType = $visioAssembly.GetType('OfficeIMO.Visio.Stencils.VisioStencilCatalog', $true)
        $measurementUnitType = $visioAssembly.GetType('OfficeIMO.Visio.VisioMeasurementUnit', $true)

        $stencil = [Activator]::CreateInstance($stencilType, [object[]]@(
            'metric.process',
            'Metric Process',
            'MetricProcess',
            'Test',
            4.0,
            2.0,
            $null,
            $null,
            $null,
            $null,
            [Enum]::Parse($measurementUnitType, 'Centimeters')))
        $stencils = [Array]::CreateInstance($stencilType, 1)
        $stencils.SetValue($stencil, 0)
        $catalog = [Activator]::CreateInstance($catalogType, [object[]]@('Metric', $stencils))

        New-OfficeVisio -Path $path -Title 'Metric stencil defaults' {
            Import-OfficeVisioStencil -Catalog $catalog -Name Metric -Default | Out-Null
            VisioPage 'Metric Page' -Width 210 -Height 297 -Unit Millimeters {
                VisioStencil -Catalog Metric -Stencil 'metric.process' -Key process -Text 'Metric' -X 50 -Y 50 -Width 30
            }
        } | Out-Null

        $document = Get-OfficeVisio -Path $path
        $shape = $document.Pages[1].Shapes[0]
        [Math]::Round($shape.Width, 3) | Should -Be ([Math]::Round(30 / 25.4, 3))
        [Math]::Round($shape.Height, 3) | Should -Be ([Math]::Round(20 / 25.4, 3))
    }

    It 'resolves relative PNG font paths from the PowerShell location' {
        $fontSource = @(
            'C:\Windows\Fonts\arial.ttf',
            '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
            '/System/Library/Fonts/Supplemental/Arial.ttf',
            '/System/Library/Fonts/Supplemental/Arial Unicode.ttf'
        ) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

        if (-not $fontSource) {
            Set-ItResult -Skipped -Because 'No TrueType font fixture was found on this runner.'
            return
        }

        $path = Join-Path $TestDrive 'font-path.vsdx'
        $pngPath = Join-Path $TestDrive 'font-path.png'
        $fontFolder = Join-Path $TestDrive 'Fonts'
        New-Item -Path $fontFolder -ItemType Directory | Out-Null
        Copy-Item -LiteralPath $fontSource -Destination (Join-Path $fontFolder 'test-font.ttf')

        New-OfficeVisio -Path $path -Title 'Font path' {
            VisioRectangle -Key text -Text 'Font Path' -X 2 -Y 2 -Width 2 -Height 1
        } | Out-Null

        Push-Location $TestDrive
        try {
            ConvertTo-OfficeVisioPng -Path $path -OutputPath $pngPath -FontFilePath '.\Fonts\test-font.ttf' | Out-Null
        } finally {
            Pop-Location
        }

        Test-Path -LiteralPath $pngPath | Should -BeTrue
    }
}
