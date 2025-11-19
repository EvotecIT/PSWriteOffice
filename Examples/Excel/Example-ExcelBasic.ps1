Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$data = @(
    [PSCustomObject]@{ Region = 'North America'; Revenue = 125000; YoY = 0.12 }
    [PSCustomObject]@{ Region = 'EMEA'; Revenue = 98000; YoY = 0.22 }
    [PSCustomObject]@{ Region = 'APAC'; Revenue = 143000; YoY = 0.18 }
)

$path = Join-Path $documents 'Excel-Basic.xlsx'

New-OfficeExcel -Path $path {
    Add-OfficeExcelSheet -Name 'Summary' -Content {
        Set-OfficeExcelCell -Address 'A1' -Value 'Region'
        Set-OfficeExcelCell -Address 'B1' -Value 'Revenue'
        Set-OfficeExcelCell -Address 'C1' -Value 'YoY'

        Add-OfficeExcelTable -Data $data -TableName 'Sales' -TableStyle 'TableStyleMedium9'
    }
} -PassThru | Out-Null

Write-Host "Workbook saved to $path"
