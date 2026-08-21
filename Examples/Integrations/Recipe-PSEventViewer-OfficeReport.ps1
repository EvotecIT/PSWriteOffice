param(
    [string] $OutputDirectory = '.',
    [string] $LogName = 'System',
    [int] $MaxEvents = 200
)

Import-Module PSEventViewer -ErrorAction Stop
Import-Module PSWriteOffice -ErrorAction Stop

$events = @(Get-EVXEvent `
    -LogName $LogName `
    -Level 1, 2, 3 `
    -TimePeriod Last24Hours `
    -ReadMode Message `
    -MaxEvents $MaxEvents)

$rows = @($events | Select-Object TimeCreated, Id, ProviderName, LevelDisplayName, MachineName, Message)
if ($rows.Count -eq 0) {
    $rows = @([pscustomobject]@{
        TimeCreated     = Get-Date
        Id              = $null
        ProviderName    = $null
        LevelDisplayName = 'Information'
        MachineName     = $env:COMPUTERNAME
        Message         = "No warning or error events were returned from $LogName in the last 24 hours."
    })
}

$excelPath = Join-Path $OutputDirectory 'Event-Report.xlsx'
$wordPath = Join-Path $OutputDirectory 'Event-Report.docx'

$rows | Export-OfficeExcel `
    -Path $excelPath `
    -WorksheetName 'Events' `
    -TableName 'EventReport' `
    -AutoFit `
    -FreezeTopRow

New-OfficeWord -Path $wordPath -Content {
    Add-OfficeWordParagraph -Text "Event report: $LogName" -Style Heading1
    Add-OfficeWordParagraph -Text "Warnings and errors returned: $($events.Count)"
    Add-OfficeWordTable -InputObject $rows -Style GridTable4Accent1 -Layout AutoFitToWindow
}

[pscustomobject]@{
    ExcelReport = Get-Item -LiteralPath $excelPath
    WordReport  = Get-Item -LiteralPath $wordPath
}
