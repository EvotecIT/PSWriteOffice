function New-OfficeWord {
    [cmdletBinding()]
    param(
        [string] $FilePath,
        [switch] $AutoSave
    )
    try {
        $WordDocument = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Create($FilePath, [DocumentFormat.OpenXml.WordprocessingDocumentType]::Document, $AutoSave.IsPresent)
    } catch {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning "New-OfficeWord - Couldn't create new Word Document at $FilePath. Error: $($_.Exception.Message)"
            return
        }
    }
    $null = $WordDocument.AddMainDocumentPart();
    $WordDocument.MainDocumentPart.Document = [DocumentFormat.OpenXml.Wordprocessing.Document]::new()
    $WordDocument.MainDocumentPart.Document.Body = [DocumentFormat.OpenXml.Wordprocessing.Body]::new()
    $WordDocument | Add-Member -Name 'FilePath' -Value $FilePath -Force -MemberType NoteProperty
    $WordDocument #.MainDocumentPart.Document
}