function New-OfficePowerPoint {
    [cmdletBinding()]
    param(
        [string] $FilePath,
        [DocumentFormat.OpenXml.PresentationDocumentType] $Type = [DocumentFormat.OpenXml.PresentationDocumentType]::Presentation #,
        #[switch] $AutoSave
    )
    <#
static DocumentFormat.OpenXml.Packaging.PresentationDocument Create(string path, DocumentFormat.OpenXml.PresentationDocumentType type)
static DocumentFormat.OpenXml.Packaging.PresentationDocument Create(System.IO.Stream stream, DocumentFormat.OpenXml.PresentationDocumentType type)
static DocumentFormat.OpenXml.Packaging.PresentationDocument Create(System.IO.Packaging.Package package, DocumentFormat.OpenXml.PresentationDocumentType type)
static DocumentFormat.OpenXml.Packaging.PresentationDocument Create(string path, DocumentFormat.OpenXml.PresentationDocumentType type, bool autoSave)
static DocumentFormat.OpenXml.Packaging.PresentationDocument Create(System.IO.Stream stream, DocumentFormat.OpenXml.PresentationDocumentType type, bool autoSave)
static DocumentFormat.OpenXml.Packaging.PresentationDocument Create(System.IO.Packaging.Package package, DocumentFormat.OpenXml.PresentationDocumentType type, bool autoSave)
#>


    $PowerPoint = [DocumentFormat.OpenXml.Packaging.PresentationDocument]::Create($FilePath, $Type, $true)
    $null = $PowerPoint.AddPresentationPart()
    $PowerPoint.PresentationPart.Presentation = [DocumentFormat.OpenXml.Presentation.Presentation]::new()
    $PowerPoint | Add-Member -Name 'FilePath' -Value $FilePath -Force -MemberType NoteProperty
    $PowerPoint
}