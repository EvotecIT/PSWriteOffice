Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$path = Join-Path $documents 'Example-WordAdvanced.docx'
$data = @(
    [pscustomobject]@{ Item = 'Alpha'; Total = 1200 }
    [pscustomobject]@{ Item = 'Beta'; Total = 800 }
    [pscustomobject]@{ Item = 'Gamma'; Total = 1500 }
)

New-OfficeWord -Path $path {
    WordSection {
        WordHeader { WordParagraph -Text 'Project Status' -Style Heading2 }
        WordFooter { WordPageNumber }

        WordParagraph -Text 'Executive Summary' -Style Heading1
        WordParagraph -Text 'This report was generated automatically.'

        WordTableOfContent -Style Template1

        WordParagraph -Text 'Status Overview' -Style Heading2
        WordList -Style Bulleted {
            WordListItem -Text 'Timeline on track'
            WordListItem -Text 'Budget variance under 5%'
        }

        WordParagraph -Text 'Metrics' -Style Heading2
        WordTable -InputObject $data -Style TableGrid {
            WordTableCondition -FilterScript { $_.Total -gt 1000 } -Style Accent1
        }

        WordParagraph -Text 'Approvals' -Style Heading2
        WordCheckBox -Title 'Approved' -Tag 'approved'
        WordDatePicker -Title 'Due Date' -Tag 'due-date'
        WordDropDownList -Title 'Status' -Tag 'status' -Items 'New','In Progress','Done'

        WordWatermark -Text 'CONFIDENTIAL'
    }
} -Open

Write-Host "Document saved to $path"
