$path = '.\Markdown-Inspection.md'
MarkdownNew -Path $path {
    MarkdownFrontMatter -Data @{ title = 'Service review'; owner = 'Operations' }
    MarkdownHeading -Level 1 -Text 'Service review'
    MarkdownParagraph -Text 'This document is inspected without regex parsing.'
    MarkdownHeading -Level 2 -Text 'Controls'
    MarkdownTable -InputObject @([pscustomobject]@{ Control = 'Backups'; Status = 'Ready' })
}

Get-OfficeMarkdownFrontMatter -InputPath $path
Get-OfficeMarkdownHeading -InputPath $path
Get-OfficeMarkdownTable -InputPath $path -AsObject
