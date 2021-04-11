function New-OfficeWord {
    [cmdletBinding()]
    param(
        [string] $FilePath
    )
    $WordDocument = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Create($FilePath, [DocumentFormat.OpenXml.WordprocessingDocumentType]::Document)
    $null = $WordDocument.AddMainDocumentPart();
    $WordDocument.MainDocumentPart.Document = [DocumentFormat.OpenXml.Wordprocessing.Document]::new()
    $WordDocument.MainDocumentPart.Document.Body = [DocumentFormat.OpenXml.Wordprocessing.Body]::new()
    $WordDocument | Add-Member -Name 'FilePath' -Value $FilePath -Force -MemberType NoteProperty
    $WordDocument #.MainDocumentPart.Document
}