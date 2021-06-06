function ConvertFrom-HTMLtoWord {
    [cmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $FilePath,
        [Parameter(Mandatory)][string] $HTML,
        [switch] $Show
    )

    $Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\TestHTML.docx

    try {
        $Converter = [HtmlToOpenXml.HtmlConverter]::new($Document.MainDocumentPart)
        $Converter.ParseHtml($HTML)
    } catch {
        Write-Warning -Message "ConvertFrom-HTMLtoWord - Couldn't convert HTML to Word. Error: $($_.Exception.Message)"
    }

    Save-OfficeWord -Document $Document -Show:$Show.IsPresent
}