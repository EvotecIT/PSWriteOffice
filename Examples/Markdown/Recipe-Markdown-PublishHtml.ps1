param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Markdown')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$markdownPath = Join-Path $OutputDirectory 'Markdown-Publish-Source.md'
$htmlPath = Join-Path $OutputDirectory 'Markdown-Published.html'
MarkdownNew -Path $markdownPath {
    MarkdownHeading -Level 1 -Text 'Operations handbook'
    MarkdownCallout -Kind warning -Title 'Before deployment' -Body 'Confirm the maintenance window.'
    MarkdownTaskList -Items 'Back up configuration','Notify users','Run health checks'
}

ConvertTo-OfficeMarkdownHtml -InputPath $markdownPath -OutputPath $htmlPath -DocumentMode `
    -Title 'Operations handbook' -IncludeAnchorLinks -ExternalLinksTargetBlank -PassThru | Out-Null

[pscustomobject]@{
    Markdown = $markdownPath
    Html     = $htmlPath
    Bytes    = (Get-Item $htmlPath).Length
} | Format-List
