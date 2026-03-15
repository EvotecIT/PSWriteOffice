Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Excel-UrlLinks.xlsx'
$rows = @(
    [PSCustomObject]@{ RFC = 'rfc7208'; Spec = 'rfc5321' }
    [PSCustomObject]@{ RFC = 'rfc7489'; Spec = 'rfc1035' }
)

New-OfficeExcel -Path $path {
    Add-OfficeExcelSheet -Name 'Summary' -Content {
        Add-OfficeExcelTable -Data $rows -TableName 'LinksTable' -AutoFit
        Set-OfficeExcelCell -Address 'D1' -Value 'Spec'
        Set-OfficeExcelCell -Address 'D2' -Value 'rfc5321'
        Set-OfficeExcelCell -Address 'D3' -Value 'rfc1035'

        Set-OfficeExcelUrlLinksByHeader -Header 'RFC' -TableName 'LinksTable' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" } -TitleScript { param($text) "Open $text" }
        Set-OfficeExcelUrlLinks -Range 'D2:D3' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" }
    }
} | Out-Null

Write-Host "Workbook saved to $path"
