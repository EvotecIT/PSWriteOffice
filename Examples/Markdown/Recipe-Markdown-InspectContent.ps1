param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Markdown')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$path = Join-Path $OutputDirectory 'Markdown-Inspection.md'
MarkdownNew -Path $path {
    MarkdownFrontMatter -Data @{ title = 'Service review'; owner = 'Operations' }
    MarkdownHeading -Level 1 -Text 'Service review'
    MarkdownParagraph -Text 'This document is inspected without regex parsing.'
    MarkdownHeading -Level 2 -Text 'Controls'
    MarkdownTable -InputObject @([pscustomobject]@{ Control = 'Backups'; Status = 'Ready' })
}

$headings = @(Get-OfficeMarkdownHeading -InputPath $path)
$frontMatter = @(Get-OfficeMarkdownFrontMatter -InputPath $path)
$tables = @(Get-OfficeMarkdownTable -InputPath $path -AsObject)
[pscustomobject]@{
    Path        = $path
    Headings    = $headings.Count
    FrontMatter = $frontMatter.Count
    TableRows   = $tables.Count
} | Format-List
