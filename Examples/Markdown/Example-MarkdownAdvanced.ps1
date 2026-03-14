Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Example-MarkdownAdvanced.md'
$data = @(
    [pscustomobject]@{ Metric = 'Latency'; Value = '120ms' }
    [pscustomobject]@{ Metric = 'Errors'; Value = '0.2%' }
)

New-OfficeMarkdown -Path $path {
    MarkdownHeading -Level 1 -Text 'Operations Report'
    MarkdownParagraph -Text 'Generated with PSWriteOffice.'
    MarkdownCallout -Kind 'note' -Title 'Reminder' -Body 'Review the incident log before publishing.'
    MarkdownHeading -Level 2 -Text 'Key Metrics'
    MarkdownTable -InputObject $data
    MarkdownCode -Language 'powershell' -Content 'Get-Service | Select-Object -First 5'
    MarkdownHorizontalRule
    MarkdownQuote -Text 'Availability is a feature.'
} -PassThru | Out-Null

Write-Host "Markdown saved to $path"
