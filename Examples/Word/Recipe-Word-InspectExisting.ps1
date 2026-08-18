param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Word')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$path = Join-Path $OutputDirectory 'Word-Inspection-Source.docx'
$services = @(
    [pscustomobject]@{ Service = 'Identity'; Owner = 'IAM'; Status = 'Ready' }
    [pscustomobject]@{ Service = 'Messaging'; Owner = 'Collaboration'; Status = 'Review' }
)

WordNew -Path $path {
    WordSection {
        WordParagraph -Text 'Service Readiness' -Style Heading1
        WordParagraph -Text 'Review the Messaging service before publication.'
        WordTable -InputObject $services -Style TableGrid
    }
}

$document = Get-OfficeWord -Path $path -ReadOnly
try {
    $paragraphs = @($document | Get-OfficeWordParagraph)
    $tables = @($document | Get-OfficeWordTable)
    $statistics = Get-OfficeWordStatistics -Document $document
    $matches = @(Find-OfficeWord -Document $document -Text 'Messaging')

    [pscustomobject]@{
        Path       = $path
        Sections   = @($document | Get-OfficeWordSection).Count
        Paragraphs = $paragraphs.Count
        Tables     = $tables.Count
        Words      = $statistics.Words
        Matches    = $matches.Count
    } | Format-List
} finally {
    Close-OfficeWord -Document $document
}
