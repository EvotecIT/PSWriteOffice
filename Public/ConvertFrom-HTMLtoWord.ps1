function ConvertFrom-HTMLtoWord {
    <#
    .SYNOPSIS
    Converts HTML input to Microsoft Word Document

    .DESCRIPTION
    Converts HTML input to Microsoft Word Document

    .PARAMETER OutputFile
    Path to the file to save converted HTML

    .PARAMETER FileHTML
    Input HTML loaded straight from file

    .PARAMETER SourceHTML
    Input HTML loaded from string

    .PARAMETER Show
    Once conversion ends show the resulting document

    .EXAMPLE
    $Objects = @(
        [PSCustomObject] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
        [PSCustomObject] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
    )

    New-HTML {
        New-HTMLText -Text 'This is a test', ' another test' -FontSize 30pt
        New-HTMLTable -DataTable $Objects -Simplify
    } -Online -FilePath $PSScriptRoot\Documents\Test.html

    ConvertFrom-HTMLToWord -OutputFile $PSScriptRoot\Documents\TestHTML.docx -FileHTML $PSScriptRoot\Documents\Test.html -Show

    .EXAMPLE
    $Objects = @(
    [PSCustomObject] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
    [PSCustomObject] @{ Test = 1; Test2 = 'Test'; Test3 = 'Ok' }
    )

    $Test = New-HTML {
        New-HTMLText -Text 'This is a test', ' another test' -FontSize 30pt
        New-HTMLTable -DataTable $Objects -simplify
    } -Online

    ConvertFrom-HTMLToWord -OutputFile $PSScriptRoot\Documents\TestHTML.docx -HTML $Test -Show

    .NOTES
    General notes
    #>
    [cmdletBinding(DefaultParameterSetName = 'HTML')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'HTMLFile')]
        [Parameter(Mandatory, ParameterSetName = 'HTML')]
        [string] $OutputFile,

        [Parameter(Mandatory, ParameterSetName = 'HTMLFile')][alias('InputFile')][string] $FileHTML,
        [Parameter(Mandatory, ParameterSetName = 'HTML')][alias('HTML')][string] $SourceHTML,

        [Parameter(Mandatory, ParameterSetName = 'HTMLFile')]
        [Parameter(Mandatory, ParameterSetName = 'HTML')]
        [switch] $Show
    )

    $Document = New-OfficeWord -FilePath $OutputFile

    if ($FileHTML) {
        $HTML = Get-Content -LiteralPath $FileHTML -Raw
    } elseif ($SourceHTML) {
        $HTML = $SourceHTML
    }

    try {
        $Converter = [HtmlToOpenXml.HtmlConverter]::new($Document._document.MainDocumentPart)
        $Converter.ParseHtml($HTML)
    } catch {
        Write-Warning -Message "ConvertFrom-HTMLtoWord - Couldn't convert HTML to Word. Error: $($_.Exception.Message)"
    }

    Save-OfficeWord -Document $Document -Show:$Show.IsPresent
}