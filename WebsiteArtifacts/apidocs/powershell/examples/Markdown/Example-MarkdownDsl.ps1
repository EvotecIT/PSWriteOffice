Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Example-MarkdownDsl.md'
$data = @(
    [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
    [pscustomobject]@{ Name = 'Beta'; Value = 2 }
)

New-OfficeMarkdown -Path $path {
    MarkdownHeading -Level 1 -Text 'Markdown DSL'
    MarkdownParagraph -Text 'Generated with PSWriteOffice.'
    MarkdownList -Items 'Alpha','Beta','Gamma'
    MarkdownTable -InputObject $data
    MarkdownCode -Language 'powershell' -Content 'Get-Date'
    MarkdownHorizontalRule
    MarkdownQuote -Text 'Ship fast, learn faster.'
} -PassThru | Out-Null

Write-Host "Markdown saved to $path"
