param(
    [switch] $Open
)

$modulePath = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
    $env:PSWRITEOFFICE_MODULE_MANIFEST
} else {
    (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1')
}
if (-not (Get-Module -Name PSWriteOffice)) { Import-Module $modulePath -ErrorAction Stop }

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Showcase-Excel-OperationalDashboard.xlsx'
$officeimoRoot = Join-Path $PSScriptRoot '..\..\..\OfficeIMO'
$logoPath = Join-Path (Join-Path $officeimoRoot 'Assets') 'OfficeIMO.png'

$services = @(
    [pscustomobject]@{ Service = 'Identity'; Owner = 'Platform'; HealthScore = 96; Incidents = 1; SlaPercent = 0.999; Status = 'Healthy'; Evidence = 'identity-runbook' }
    [pscustomobject]@{ Service = 'Mail Flow'; Owner = 'Messaging'; HealthScore = 89; Incidents = 4; SlaPercent = 0.995; Status = 'Watch'; Evidence = 'mail-flow-kpis' }
    [pscustomobject]@{ Service = 'Endpoint Fleet'; Owner = 'Workplace'; HealthScore = 82; Incidents = 7; SlaPercent = 0.991; Status = 'Watch'; Evidence = 'endpoint-drift' }
    [pscustomobject]@{ Service = 'Backup'; Owner = 'Operations'; HealthScore = 74; Incidents = 11; SlaPercent = 0.986; Status = 'Risk'; Evidence = 'backup-exceptions' }
    [pscustomobject]@{ Service = 'Network Edge'; Owner = 'Network'; HealthScore = 93; Incidents = 2; SlaPercent = 0.997; Status = 'Healthy'; Evidence = 'edge-availability' }
    [pscustomobject]@{ Service = 'Data Warehouse'; Owner = 'Data'; HealthScore = 78; Incidents = 8; SlaPercent = 0.988; Status = 'Risk'; Evidence = 'warehouse-latency' }
    [pscustomobject]@{ Service = 'Collaboration'; Owner = 'Productivity'; HealthScore = 91; Incidents = 3; SlaPercent = 0.996; Status = 'Healthy'; Evidence = 'collab-telemetry' }
    [pscustomobject]@{ Service = 'Security Signals'; Owner = 'Security'; HealthScore = 87; Incidents = 5; SlaPercent = 0.993; Status = 'Watch'; Evidence = 'signal-coverage' }
)

$trend = @(
    [pscustomobject]@{ Month = 'Jan'; Availability = 99.1; Incidents = 14; Automation = 62 }
    [pscustomobject]@{ Month = 'Feb'; Availability = 99.2; Incidents = 12; Automation = 66 }
    [pscustomobject]@{ Month = 'Mar'; Availability = 99.4; Incidents = 10; Automation = 71 }
    [pscustomobject]@{ Month = 'Apr'; Availability = 99.3; Incidents = 11; Automation = 74 }
    [pscustomobject]@{ Month = 'May'; Availability = 99.6; Incidents = 8; Automation = 79 }
    [pscustomobject]@{ Month = 'Jun'; Availability = 99.7; Incidents = 6; Automation = 83 }
)

$legend = @(
    [pscustomobject]@{ Status = 'Healthy'; Meaning = 'Stable service posture'; Action = 'Keep monitoring' }
    [pscustomobject]@{ Status = 'Watch'; Meaning = 'Needs routine follow-up'; Action = 'Review trend and owners' }
    [pscustomobject]@{ Status = 'Risk'; Meaning = 'Needs active remediation'; Action = 'Open mitigation plan' }
)

$ownerSummary = $services |
    Group-Object Owner |
    ForEach-Object {
        [pscustomobject]@{
            Owner         = $_.Name
            Services      = $_.Count
            AverageHealth = [math]::Round(($_.Group | Measure-Object HealthScore -Average).Average, 1)
            Incidents     = ($_.Group | Measure-Object Incidents -Sum).Sum
            RiskServices  = @($_.Group | Where-Object Status -eq 'Risk').Count
        }
    }

New-OfficeExcel -Path $path {
    ExcelSheet 'Summary' {
        ExcelGridlines -Hide
        ExcelCell -Address 'A1' -Value 'Operational Dashboard'
        ExcelCell -Address 'A2' -Value 'Generated with PSWriteOffice from PowerShell objects.'
        ExcelCell -Address 'A4' -Value 'Average Health'
        ExcelCell -Address 'B4' -Formula 'AVERAGE(Services!C2:C9)' -NumberFormat '0.0'
        ExcelCell -Address 'D4' -Value 'Total Incidents'
        ExcelCell -Address 'E4' -Formula 'SUM(Services!D2:D9)' -NumberFormat '0'
        ExcelCell -Address 'G4' -Value 'Risk Services'
        ExcelCell -Address 'H4' -Formula 'COUNTIF(Services!F2:F9,"Risk")' -NumberFormat '0'

        ExcelTable -Data $legend -TableName 'StatusLegend' -StartRow 7 -StartColumn 1 -TableStyle 'TableStyleMedium4' -AutoFit
        ExcelChart -Range 'A7:B10' -Row 7 -Column 6 -Type Doughnut -Title 'Status Meaning Mix' -WidthPixels 440 -HeightPixels 260 |
            Set-OfficeExcelChartLegend -Position Right |
            Set-OfficeExcelChartDataLabels -ShowValue $true -ShowCategoryName $true -Position OutsideEnd |
            Set-OfficeExcelChartStyle -StyleId 251 -ColorStyleId 10

        if (Test-Path $logoPath) {
            ExcelImage -Path $logoPath -Address 'J1' -WidthPixels 140 -HeightPixels 52 | Out-Null
        }

        ExcelHeaderFooter -HeaderCenter 'PSWriteOffice operational dashboard' -FooterRight 'Page &P of &N'
        ExcelOrientation -Orientation Landscape
        ExcelMargins -Left 0.4 -Right 0.4 -Top 0.7 -Bottom 0.7
        ExcelPageSetup -FitToWidth 1 -FitToHeight 0
    }

    ExcelSheet 'Services' {
        ExcelTable -Data $services -TableName 'ServiceHealth' -StartRow 1 -StartColumn 1 -TableStyle 'TableStyleMedium9' -AutoFit
        ExcelFreeze -TopRows 1
        ExcelValidationList -Range 'F2:F50' -Values 'Healthy','Watch','Risk'
        ExcelConditionalColorScale -Range 'C2:C9' -StartColor '#F8696B' -EndColor '#63BE7B'
        ExcelConditionalDataBar -Range 'D2:D9' -Color '#5B9BD5'
        ExcelConditionalIconSet -Range 'C2:C9' -IconSet ThreeTrafficLights1 -Reverse $true
        ExcelUrlLinksByHeader -Header 'Evidence' -TableName 'ServiceHealth' -UrlScript { param($text) "https://evotec.xyz/docs/$text" } -TitleScript { param($text) "Open $text" }
        ExcelChart -TableName 'ServiceHealth' -Row 12 -Column 1 -Type BarClustered -Title 'Health Score and Incidents' -WidthPixels 760 -HeightPixels 340 |
            Set-OfficeExcelChartLegend -Position Bottom |
            Set-OfficeExcelChartDataLabels -ShowValue $true -Position OutsideEnd |
            Set-OfficeExcelChartStyle -StyleId 251 -ColorStyleId 10
        ExcelHeaderFooter -HeaderCenter 'Service details' -FooterRight 'Page &P of &N'
    }

    ExcelSheet 'Trend' {
        ExcelTable -Data $trend -TableName 'TrendData' -StartRow 1 -StartColumn 1 -TableStyle 'TableStyleMedium2' -AutoFit
        ExcelFreeze -TopRows 1
        ExcelChart -TableName 'TrendData' -Row 10 -Column 1 -Type Line -Title 'Availability, Incidents, and Automation' -WidthPixels 780 -HeightPixels 340 |
            Set-OfficeExcelChartLegend -Position Bottom |
            Set-OfficeExcelChartDataLabels -ShowValue $true -Position Top |
            Set-OfficeExcelChartStyle -StyleId 251 -ColorStyleId 10
        ExcelHeaderFooter -HeaderCenter 'Trend and automation' -FooterRight 'Page &P of &N'
    }

    ExcelSheet 'Owner Summary' {
        ExcelTable -Data $ownerSummary -TableName 'OwnerSummary' -StartRow 1 -StartColumn 1 -TableStyle 'TableStyleMedium5' -AutoFit
        ExcelConditionalDataBar -Range 'D2:D20' -Color '#ED7D31'
        ExcelConditionalIconSet -Range 'C2:C20' -IconSet ThreeTrafficLights1 -Reverse $true
        ExcelHeaderFooter -HeaderCenter 'Owner summary' -FooterRight 'Page &P of &N'
    }

    ExcelSheet 'Notes' {
        ExcelCell -Address 'A1' -Value 'Generation Notes'
        ExcelCell -Address 'A2' -Value 'This sheet is hidden and carries inputs useful for audit/debugging.'
        ExcelCell -Address 'A4' -Value 'Created'
        ExcelCell -Address 'B4' -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        ExcelCell -Address 'A5' -Value 'Source'
        ExcelCell -Address 'B5' -Value 'Examples/Showcase/Showcase-Excel-OperationalDashboard.ps1'
        ExcelSheetVisibility -Hide
    }

    ExcelTableOfContents -SheetName 'Index' -IncludeNamedRanges -AddBackLinks -BackLinkText 'Back to Index'
} -Open:$Open

$summary = Get-OfficeExcelSummary -Path $path
Write-Host "Workbook saved to $path"
Write-Host "Workbook summary: $($summary.SheetCount) sheets, $($summary.TableCount) tables, $($summary.ChartCount) charts, $($summary.HyperlinkCount) links"
