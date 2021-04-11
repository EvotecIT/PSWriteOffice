<#
function New-OfficeExcel {
    [cmdletBinding()]
    param(
        [string] $FilePath,
        [DocumentFormat.OpenXml.SpreadsheetDocumentType] $Type = [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook,
        [switch] $Template
    )
}
#>