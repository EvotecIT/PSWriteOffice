function New-OfficeWord {
    [cmdletBinding()]
    param(
        [string] $FilePath,
        [switch] $AutoSave,
        [int] $Retry = 2
    )
    $Saved = $false
    $Count = 0
    while ($Count -le $Retry -and $Saved -eq $false) {
        $Count++
        try {
            $WordDocument = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Create($FilePath, [DocumentFormat.OpenXml.WordprocessingDocumentType]::Document, $AutoSave.IsPresent)
            $Saved = $true
        } catch {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw
            } else {
                Write-Warning "New-OfficeWord - Couldn't create new Word Document at $FilePath. Error: $($_.Exception.Message)"
            }
        }
        if (-not $Saved) {
            if ($Retry -ge $Count) {
                $FilePath = [io.path]::GetTempFileName().Replace('.tmp', '.docx')
                Write-Warning -Message "New-OfficeWord - Couldn't save using provided file name, retrying with $FilePath"
            } else {
                Write-Warning -Message "New-OfficeWord - Couldn't save using provided file name. Run out of retries ($Count / $Retry)."
                return
            }
        }
    }
    $null = $WordDocument.AddMainDocumentPart();
    $WordDocument.MainDocumentPart.Document = [DocumentFormat.OpenXml.Wordprocessing.Document]::new()
    $WordDocument.MainDocumentPart.Document.Body = [DocumentFormat.OpenXml.Wordprocessing.Body]::new()
    $WordDocument | Add-Member -Name 'FilePath' -Value $FilePath -Force -MemberType NoteProperty
    $WordDocument #.MainDocumentPart.Document
}