$ErrorActionPreference = 'Stop'

$modulePath = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
    $env:PSWRITEOFFICE_MODULE_MANIFEST
} else {
    Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1'
}

if (-not (Get-Module -Name PSWriteOffice)) {
    Import-Module $modulePath -ErrorAction Stop
}

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Showcase-Word-ExecutiveReport.docx'
$heroPath = Join-Path $documents 'Showcase-Word-ExecutiveReport-Hero.png'
Remove-Item -Path $path -Force -ErrorAction SilentlyContinue
Remove-Item -Path $heroPath -Force -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.Drawing
$bitmap = [System.Drawing.Bitmap]::new(960, 220)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$accentBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::Teal)
$darkBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::DarkSlateGray)
$whiteBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
$inkBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(34, 42, 53))
$mutedBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(84, 96, 111))
try {
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.Clear([System.Drawing.Color]::FromArgb(246, 249, 251))

    $graphics.FillRectangle($accentBrush, 0, 0, 18, 220)
    $graphics.FillRectangle($darkBrush, 18, 0, 140, 220)
    $graphics.FillEllipse($whiteBrush, 52, 42, 72, 72)
    $graphics.FillEllipse($accentBrush, 70, 60, 36, 36)

    $titleFont = [System.Drawing.Font]::new('Segoe UI', 32, [System.Drawing.FontStyle]::Bold)
    $subtitleFont = [System.Drawing.Font]::new('Segoe UI', 16, [System.Drawing.FontStyle]::Regular)
    $labelFont = [System.Drawing.Font]::new('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $graphics.DrawString('Executive Service Health Report', $titleFont, $inkBrush, 190, 48)
    $graphics.DrawString('Generated from PowerShell objects with PSWriteOffice and OfficeIMO', $subtitleFont, $mutedBrush, 194, 102)
    $graphics.DrawString('TOC   TABLES   CHARTS   NOTES   APPROVALS', $labelFont, $accentBrush, 196, 154)
} finally {
    if ($titleFont) { $titleFont.Dispose() }
    if ($subtitleFont) { $subtitleFont.Dispose() }
    if ($labelFont) { $labelFont.Dispose() }
    $accentBrush.Dispose()
    $darkBrush.Dispose()
    $whiteBrush.Dispose()
    $inkBrush.Dispose()
    $mutedBrush.Dispose()
    $graphics.Dispose()
    $bitmap.Save($heroPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
}

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
        WordImage -Path $heroPath -Width 640 -Height 147 -Description 'Executive service health report banner'
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

        WordTableOfContent -Style Template1

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
        Update-OfficeWordTableOfContent
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
