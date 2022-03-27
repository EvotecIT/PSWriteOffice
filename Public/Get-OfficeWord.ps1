function Get-OfficeWord {
    [cmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $FilePath,
        [switch] $ReadOnly,
        [switch] $AutoSave
    )

    if ($FilePath -and (Test-Path -LiteralPath $FilePath)) {
        try {
            [OfficeIMO.Word.WordDocument]::Load($FilePath, $ReadOnly.IsPresent, $AutoSave.IsPresent)
        } catch {
            if ($PSBoundParameters.ErrorAction -eq 'Stop') {
                throw
            } else {
                Write-Warning "Get-OfficeWord - File $FilePath couldn't be open. Error: $($_.Exception.Message)"
            }
        }
    } else {
        if ($PSBoundParameters.ErrorAction -eq 'Stop') {
            throw "File $FilePath doesn't exists. Try again."
        } else {
            Write-Warning "Get-OfficeWord - File $FilePath doesn't exists. Try again."
        }
    }

}