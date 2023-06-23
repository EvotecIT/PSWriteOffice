#Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Documents = Get-ChildItem -Filter *.docx -Path $PSScriptRoot\Documents
$Found = foreach ($File in $Documents) {
    $Document = Get-OfficeWord -FilePath $File.FullName
    $FoundWords = $document.Find("Test");
    [PSCustomObject] @{
        File  = $File.Name
        Found = $FoundWords.Count
    }
    $Document.Dispose()
}
$Found | Format-Table