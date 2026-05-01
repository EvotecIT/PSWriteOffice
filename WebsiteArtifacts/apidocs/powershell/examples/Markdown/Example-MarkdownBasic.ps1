$modulePath = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
    $env:PSWRITEOFFICE_MODULE_MANIFEST
} else {
    (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1')
}
if (-not (Get-Module -Name PSWriteOffice)) { Import-Module $modulePath -ErrorAction Stop }
$markdown = @(
    [PSCustomObject]@{ Name = 'Alpha'; Value = 1 }
    [PSCustomObject]@{ Name = 'Beta'; Value = 2 }
) | ConvertTo-OfficeMarkdown

$html = ConvertTo-OfficeMarkdownHtml -Text $markdown -DocumentMode

Write-Host "Markdown:"
Write-Host $markdown
Write-Host "HTML:"
Write-Host $html
