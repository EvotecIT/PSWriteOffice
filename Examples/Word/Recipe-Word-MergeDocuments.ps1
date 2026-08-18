param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Word')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$cover = Join-Path $OutputDirectory 'Word-Pack-Cover.docx'
$detail = Join-Path $OutputDirectory 'Word-Pack-Detail.docx'
$appendix = Join-Path $OutputDirectory 'Word-Pack-Appendix.docx'
$merged = Join-Path $OutputDirectory 'Word-Combined-Pack.docx'

WordNew -Path $cover { WordSection { WordParagraph -Text 'Operations Pack' -Style Heading1 } }
WordNew -Path $detail { WordSection { WordParagraph -Text 'Current Status' -Style Heading1; WordParagraph -Text 'All core services are available.' } }
WordNew -Path $appendix { WordSection { WordParagraph -Text 'Appendix' -Style Heading1; WordParagraph -Text 'Evidence retained for 90 days.' } }

Join-OfficeWordDocument -InputPath $cover -AppendPath $detail,$appendix -OutputPath $merged

[pscustomobject]@{
    Path       = $merged
    Paragraphs = (Get-OfficeWordStatistics -Path $merged).Paragraphs
    HasStatus  = @(Find-OfficeWord -Path $merged -Text 'Current Status').Count -gt 0
    HasAppendix = @(Find-OfficeWord -Path $merged -Text 'Appendix').Count -gt 0
} | Format-List
