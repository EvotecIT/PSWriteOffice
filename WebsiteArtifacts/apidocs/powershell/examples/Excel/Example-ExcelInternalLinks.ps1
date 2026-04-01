Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Excel-InternalLinks.xlsx'
$rows = @(
    [PSCustomObject]@{ Sheet = 'Alpha'; Target = 'Alpha' }
    [PSCustomObject]@{ Sheet = 'Beta'; Target = 'Beta' }
)

New-OfficeExcel -Path $path {
    Add-OfficeExcelSheet -Name 'Summary' -Content {
        Add-OfficeExcelTable -Data $rows -TableName 'SummaryTable' -AutoFit
        Set-OfficeExcelCell -Address 'D1' -Value 'Sheet'
        Set-OfficeExcelCell -Address 'D2' -Value 'Alpha'
        Set-OfficeExcelCell -Address 'D3' -Value 'Beta'

        Set-OfficeExcelInternalLinks -Range 'D2:D3'
        Set-OfficeExcelInternalLinksByHeader -Header 'Sheet' -TableName 'SummaryTable' -DisplayScript { param($text) "Open $text" }
    }
    Add-OfficeExcelSheet -Name 'Alpha' -Content {
        Set-OfficeExcelCell -Address 'A1' -Value 'Alpha Home'
    }
    Add-OfficeExcelSheet -Name 'Beta' -Content {
        Set-OfficeExcelCell -Address 'A1' -Value 'Beta Home'
    }
} | Out-Null

Write-Host "Workbook saved to $path"
