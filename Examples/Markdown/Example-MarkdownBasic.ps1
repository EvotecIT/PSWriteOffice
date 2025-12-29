Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$markdown = @(
    [PSCustomObject]@{ Name = 'Alpha'; Value = 1 }
    [PSCustomObject]@{ Name = 'Beta'; Value = 2 }
) | ConvertTo-OfficeMarkdown

$html = ConvertTo-OfficeMarkdownHtml -Text $markdown -DocumentMode

Write-Host "Markdown:"
Write-Host $markdown
Write-Host "HTML:"
Write-Host $html
