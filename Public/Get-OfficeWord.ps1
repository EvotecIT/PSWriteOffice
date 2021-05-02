function Get-OfficeWord {
    [cmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $FilePath,
        [switch] $ReadOnly,
        [switch] $AutoSave
    )

    if ($FilePath -and (Test-Path -LiteralPath $FilePath)) {
        $Settings = [DocumentFormat.OpenXml.Packaging.OpenSettings]::new()
        $Settings.AutoSave = $AutoSave.IsPresent
        <#
        $Settings.MarkupCompatibilityProcessSettings
        $Settings.MaxCharactersInPart
        $Settings.RelationshipErrorHandlerFactory
        $Settings
        #>

        # Byte Array and Memory Stream are the only way to make sure the original file is not overwritten
        [byte[]] $ByteArray = [System.IO.File]::ReadAllBytes($FilePath);
        $MemoryStream = [System.IO.MemoryStream]::new($ByteArray)

        try {
            #$WordDocument = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($FilePath, -not $ReadOnly.IsPresent, $Settings) #Create($FilePath, [DocumentFormat.OpenXml.WordprocessingDocumentType]::Document)
            $WordDocument = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($MemoryStream, -not $ReadOnly.IsPresent, $Settings) #Create($FilePath, [DocumentFormat.OpenXml.WordprocessingDocumentType]::Document)
            $WordDocument | Add-Member -MemberType NoteProperty -Name 'FilePath' -Value $FilePath -Force
            #$WordDocument | Add-Member -MemberType NoteProperty -Name 'MemoryStream' -Value $MemoryStream -Force
            $WordDocument
        } catch {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw
            } else {
                Write-Warning "Get-OfficeWord - File $FilePath couldn't be open. Error: $($_.Exception.Message)"
            }
        }
    } else {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw
        } else {
            Write-Warning "Get-OfficeWord - File $FilePath doesn't exists. Try again."
        }
    }

}