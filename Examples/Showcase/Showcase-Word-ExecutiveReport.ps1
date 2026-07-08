$ErrorActionPreference = 'Stop'

Import-Module PSWriteOffice -ErrorAction Stop

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Showcase-Word-ExecutiveReport.docx'
Remove-Item -Path $path -Force -ErrorAction SilentlyContinue

$services = @(
    [pscustomobject]@{ Service = 'Identity Sync'; Owner = 'Platform'; Status = 'Healthy'; Availability = 99.98; Incidents = 1; Risk = 'Low'; NextAction = 'Keep monitoring password hash sync drift' }
    [pscustomobject]@{ Service = 'Messaging'; Owner = 'Collaboration'; Status = 'Watch'; Availability = 99.72; Incidents = 4; Risk = 'Medium'; NextAction = 'Reduce transport queue alert noise' }
    [pscustomobject]@{ Service = 'File Services'; Owner = 'Core IT'; Status = 'Healthy'; Availability = 99.91; Incidents = 2; Risk = 'Low'; NextAction = 'Complete stale share owner review' }
    [pscustomobject]@{ Service = 'Remote Access'; Owner = 'Security'; Status = 'Needs action'; Availability = 98.84; Incidents = 7; Risk = 'High'; NextAction = 'Rotate legacy VPN profiles and publish owner runbook' }
    [pscustomobject]@{ Service = 'Endpoint Backup'; Owner = 'Operations'; Status = 'Watch'; Availability = 99.21; Incidents = 5; Risk = 'Medium'; NextAction = 'Tune failed backup retry window' }
)

$trend = @(
    [pscustomobject]@{ Month = 'Jan'; Availability = 99.20; Incidents = 8 }
    [pscustomobject]@{ Month = 'Feb'; Availability = 99.34; Incidents = 7 }
    [pscustomobject]@{ Month = 'Mar'; Availability = 99.55; Incidents = 6 }
    [pscustomobject]@{ Month = 'Apr'; Availability = 99.61; Incidents = 5 }
    [pscustomobject]@{ Month = 'May'; Availability = 99.73; Incidents = 4 }
    [pscustomobject]@{ Month = 'Jun'; Availability = 99.82; Incidents = 3 }
)

$executiveSignals = @(
    [pscustomobject]@{ Signal = 'Audience'; Detail = 'Technology leadership, service owners, and operational reviewers' }
    [pscustomobject]@{ Signal = 'Decision'; Detail = 'Approve the focused remediation plan for high-friction services' }
    [pscustomobject]@{ Signal = 'Evidence'; Detail = 'TOC, scorecard table, trend chart, links, notes, and approval controls' }
)

$actions = @(
    [pscustomobject]@{ Priority = 'P1'; Action = 'Retire legacy VPN profiles'; Owner = 'Security'; Due = '2026-05-24'; Outcome = 'Lower high-risk remote-access incidents' }
    [pscustomobject]@{ Priority = 'P2'; Action = 'Tune backup retry window'; Owner = 'Operations'; Due = '2026-05-31'; Outcome = 'Improve endpoint backup completion rate' }
    [pscustomobject]@{ Priority = 'P2'; Action = 'Finalize share owner review'; Owner = 'Core IT'; Due = '2026-06-07'; Outcome = 'Reduce stale access paths before audit' }
)

New-OfficeWord -Path $path {
    WordSection {
        WordHeader {
            WordParagraph {
                WordBold 'PSWriteOffice Showcase'
                WordText ' | Executive service health'
            }
        }
        WordFooter {
            WordText "Generated $(Get-Date -Format 'yyyy-MM-dd')"
            WordText ' | Page '
            WordPageNumber
        }

        Set-OfficeWordDocumentProperty -Name Title -Value 'Executive Service Health Report'
        Set-OfficeWordDocumentProperty -Name Creator -Value 'PSWriteOffice'
        Set-OfficeWordDocumentProperty -Name Category -Value 'Showcase'
        Set-OfficeWordDocumentProperty -Name ShowcaseProduct -Value 'Word' -Custom

        WordParagraph -Text 'Executive Service Health Report' -Style Heading1
        WordParagraph 'Generated from PowerShell objects with PSWriteOffice and OfficeIMO.'
        WordTable -InputObject $executiveSignals -Style GridTable5DarkAccent1 -Layout AutoFitToWindow
        WordParagraph {
            WordBold 'Audience: '
            WordText 'technology leadership, service owners, and operational reviewers.'
            WordFootnote 'This sample uses synthetic service-health data generated entirely in PowerShell.'
        }
        WordParagraph {
            WordBold 'Decision needed: '
            WordText 'approve the focused remediation plan for high-friction services and keep monthly trend review cadence.'
            WordBookmark -Name 'ExecutiveSummary'
        }

        WordTableOfContents -Style Template1

        WordParagraph -Text 'Executive Summary' -Style Heading1
        WordParagraph 'The portfolio is stable overall, but Remote Access and Endpoint Backup need targeted work before the next governance review.'
        WordList -Style Bulleted {
            WordListItem -Text 'Average availability across tracked services is above 99.5%.'
            WordListItem -Text 'Remote Access is the only high-risk service and owns the most urgent remediation action.'
            WordListItem -Text 'The next review should focus on stale access paths, retry windows, and alert quality.'
        }

        WordParagraph -Text 'Service Scorecard' -Style Heading1
        WordParagraph {
            WordText 'Rows needing attention are highlighted so the table remains useful after export.'
            WordEndnote 'Risk labels combine current incident count, owner feedback, and observed availability trend.'
        }
        WordTable -InputObject $services -Style GridTable4Accent1 -Layout AutoFitToWindow {
            WordTableCondition -FilterScript { $_.Status -eq 'Needs action' } -BackgroundColor '#fde2e2'
            WordTableCondition -FilterScript { $_.Status -eq 'Watch' } -BackgroundColor '#fff4cc'
        }

        WordParagraph -Text 'Trend and Incidents' -Style Heading1
        WordParagraph 'The line chart gives leadership a quick visual read on improvement without opening a separate workbook.'
        WordChart -Type Line -Data $trend -CategoryProperty Month -SeriesProperty Availability, Incidents -Title 'Availability and incident trend' -Legend -LegendPosition Bottom -XAxisTitle 'Month' -YAxisTitle 'Value' -FitToPageWidth -WidthFraction 0.92

        WordParagraph -Text 'Action Plan' -Style Heading1
        WordTable -InputObject $actions -Style GridTable5DarkAccent1 -Layout AutoFitToWindow
        WordParagraph {
            WordText 'Jump back to '
            WordHyperlink -Text 'Executive Summary' -Anchor 'ExecutiveSummary' -Styled
            WordText ' when reviewing the action plan.'
        }

        WordParagraph -Text 'Approval Controls' -Style Heading1
        WordParagraph {
            WordText 'Approved for publication: '
            WordCheckBox -Alias 'ApprovedForPublication' -Tag 'approval-publish'
        }
        WordParagraph {
            WordText 'Next review date: '
            WordDatePicker -Date (Get-Date '2026-06-15') -Alias 'NextReviewDate' -Tag 'review-date'
        }
        WordParagraph {
            WordText 'Review status: '
            WordDropDownList -Items 'Draft','Ready for review','Approved' -Alias 'ReviewStatus' -Tag 'review-status'
        }

        WordWatermark -Text 'SHOWCASE'
        Update-OfficeWordFields
        Update-OfficeWordTableOfContents
    }
} | Out-Null

$document = Get-OfficeWord -Path $path -ReadOnly
try {
    [pscustomobject]@{
        Path            = $path
        Paragraphs      = $document.Paragraphs.Count
        Tables          = $document.Tables.Count
        Charts          = $document.Charts.Count
        ContentControls = $document.StructuredDocumentTags.Count
        Footnotes       = @(Get-OfficeWordFootnote -Document $document).Count
        Endnotes        = @(Get-OfficeWordEndnote -Document $document).Count
    } | Format-List
} finally {
    $document.Dispose()
}
