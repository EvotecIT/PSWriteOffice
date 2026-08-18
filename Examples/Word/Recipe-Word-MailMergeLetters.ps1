param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\..\Artefacts\Examples\Word')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop
New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null

$recipients = @(
    @{ FirstName = 'Ada'; OrderId = 'SO-1042'; DeliveryDate = '2026-09-01' }
    @{ FirstName = 'Grace'; OrderId = 'SO-1043'; DeliveryDate = '2026-09-03' }
)

foreach ($recipient in $recipients) {
    $path = Join-Path $OutputDirectory ("Order-{0}.docx" -f $recipient.OrderId)
    WordNew -Path $path {
        WordSection {
            WordParagraph -Text 'Order confirmation' -Style Heading1
            WordParagraph {
                WordText 'Hello '
                WordField -Type MergeField -Parameters '"FirstName"'
                WordText ','
            }
            WordParagraph {
                WordText 'Order '
                WordField -Type MergeField -Parameters '"OrderId"'
                WordText ' is scheduled for '
                WordField -Type MergeField -Parameters '"DeliveryDate"'
                WordText '.'
            }
            Invoke-OfficeWordMailMerge -Values $recipient
        }
    }

    [pscustomobject]@{
        Path      = $path
        Recipient = $recipient.FirstName
        Verified  = @(Find-OfficeWord -Path $path -Text $recipient.OrderId).Count -gt 0
    }
}
