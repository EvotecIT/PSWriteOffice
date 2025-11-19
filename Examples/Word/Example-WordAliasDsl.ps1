Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$orders = @(
    [PSCustomObject]@{ Customer = 'Contoso'; Total = 1850; Status = 'Open' }
    [PSCustomObject]@{ Customer = 'Fabrikam'; Total = 640; Status = 'Closed' }
    [PSCustomObject]@{ Customer = 'Northwind'; Total = 2240; Status = 'Open' }
)

$docPath = Join-Path $documents 'Word-AliasDsl.docx'

New-OfficeWord -Path $docPath {
    WordSection {
        WordHeader {
            WordParagraph 'Alias DSL Sample' -Style Heading2
        }
        WordFooter {
            WordPageNumber -IncludeTotalPages
        }

        WordParagraph {
            WordText 'Hello '
            WordBold 'world'
            WordText '! This section uses pure aliases.'
        }

        WordList -Type Numbered {
            WordListItem 'Capture highlights'
            WordListItem 'Summarize blockers'
            WordListItem 'Confirm next steps'
        }

        WordTable -Data $orders -Style 'GridTable1LightAccent1' {
            WordTableCondition -FilterScript { $_.Status -eq 'Open' } -BackgroundColor '#fff4d6'
        }

        WordParagraph {
            WordText 'Generated on '
            WordBold (Get-Date -Format 'yyyy-MM-dd HH:mm')
        }
    }
} -PassThru | Out-Null

Write-Host "Document saved to $docPath"
