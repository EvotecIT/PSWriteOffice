param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Word')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$path = Join-Path $OutputDirectory 'Word-Updated-In-Place.docx'
WordNew -Path $path {
    WordSection {
        WordParagraph -Text 'FY24 Service Review' -Style Heading1
        WordParagraph {
            WordText 'Open the '
            WordHyperlink -Text 'FY24 portal' -Url 'https://reports.example.test/FY24' -Tooltip 'FY24 reports'
            WordText ' for supporting evidence.'
        }
        WordParagraph {
            WordText 'Summary'
            WordBookmark -Name 'FY24Summary'
        }
    }
}

$changes = Update-OfficeWordText -Path $path -OldValue 'FY24' -NewValue 'FY25' `
    -IncludeHyperlinkText -IncludeHyperlinkUri -IncludeHyperlinkTooltip -IncludeHyperlinkAnchor

[pscustomobject]@{
    Path          = $path
    Changes       = $changes
    UpdatedText   = @(Find-OfficeWord -Path $path -Text 'FY25').Count
    RemainingText = @(Find-OfficeWord -Path $path -Text 'FY24').Count
} | Format-List
