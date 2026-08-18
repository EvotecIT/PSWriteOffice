param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Markdown')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$markdownPath = Join-Path $OutputDirectory 'Markdown-Word-Source.md'
$wordPath = Join-Path $OutputDirectory 'Markdown-Converted.docx'
$roundTripPath = Join-Path $OutputDirectory 'Markdown-Word-RoundTrip.md'
MarkdownNew -Path $markdownPath {
    MarkdownHeading -Level 1 -Text 'Change plan'
    MarkdownParagraph -Text 'The same source can be reviewed in Word and returned to Markdown.'
    MarkdownList -Items 'Prepare','Approve','Deploy'
}

ConvertFrom-OfficeWordMarkdown -FilePath $markdownPath -OutputPath $wordPath
ConvertTo-OfficeWordMarkdown -FilePath $wordPath -OutputPath $roundTripPath

[pscustomobject]@{
    Markdown  = $markdownPath
    Word      = $wordPath
    RoundTrip = $roundTripPath
    Verified  = (Get-Content -Path $roundTripPath -Raw) -match 'Change plan'
} | Format-List
