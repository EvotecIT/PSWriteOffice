param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Excel')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$path = Join-Path $OutputDirectory 'Budget-Dashboard.xlsx'
$budget = @(
    [pscustomobject]@{ Department = 'Engineering'; Budget = 240000; Actual = 218500; Forecast = 236000 }
    [pscustomobject]@{ Department = 'Operations'; Budget = 160000; Actual = 151200; Forecast = 164000 }
    [pscustomobject]@{ Department = 'Sales'; Budget = 190000; Actual = 177800; Forecast = 188500 }
    [pscustomobject]@{ Department = 'Support'; Budget = 110000; Actual = 104300; Forecast = 109000 }
)

ExcelNew -Path $path {
    ExcelSheet 'Dashboard' {
        ExcelGridlines -Hide
        ExcelCell -Address A1 -Value 'Department Budget Dashboard'
        ExcelCell -Address A3 -Value 'Total budget'
        ExcelCell -Address B3 -Formula 'SUM(Detail!B2:B5)' -NumberFormat '$#,##0'
        ExcelCell -Address D3 -Value 'Actual spend'
        ExcelCell -Address E3 -Formula 'SUM(Detail!C2:C5)' -NumberFormat '$#,##0'
        ExcelCell -Address G3 -Value 'Forecast variance'
        ExcelCell -Address H3 -Formula 'SUM(Detail!D2:D5)-SUM(Detail!B2:B5)' -NumberFormat '$#,##0;[Red]-$#,##0'
        ExcelOrientation -Orientation Landscape
        ExcelPageSetup -FitToWidth 1 -FitToHeight 0
    }

    ExcelSheet 'Detail' {
        ExcelTable -Data $budget -TableName 'DepartmentBudget' -StartRow 1 -StartColumn 1 -TableStyle 'TableStyleMedium4' -AutoFit
        ExcelFreeze -TopRows 1
        ExcelConditionalDataBar -Range 'C2:C5' -Color '#5B9BD5'
        ExcelConditionalIconSet -Range 'D2:D5' -IconSet ThreeTrafficLights1
        foreach ($header in 'Budget', 'Actual', 'Forecast') {
            ExcelColumnStyleByHeader -Header $header -NumberFormat '$#,##0' -AutoFit
        }
        ExcelChart -Range 'A1:D5' -Row 7 -Column 1 -Type ColumnClustered -Title 'Budget, actual, and forecast' -WidthPixels 780 -HeightPixels 340
    }

    ExcelTableOfContents -SheetName 'Index' -AddBackLinks -BackLinkText 'Back to Index'
}

Write-Host "Excel budget dashboard saved to $path"
