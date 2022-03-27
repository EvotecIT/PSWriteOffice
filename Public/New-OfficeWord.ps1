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
            $WordDocument = [OfficeImo.Word.WordDocument]::Create($FilePath, $AutoSave.IsPresent)
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
                $FilePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "$($([System.IO.Path]::GetRandomFileName()).Split('.')[0]).docx")
                Write-Warning -Message "New-OfficeWord - Couldn't save using provided file name, retrying with $FilePath"
            } else {
                Write-Warning -Message "New-OfficeWord - Couldn't save using provided file name. Run out of retries ($Count / $Retry)."
                return
            }
        }
    }
    $WordDocument
}