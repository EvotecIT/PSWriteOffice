# This is just a show what can be quickly done using .NET before I get to do it's PowerShell version

Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

$Document = New-OfficeWord -FilePath $PSScriptRoot\Documents\BasicDocumentWithHyperlinks.docx
$Document.BuiltinDocumentProperties.Title = "This is title"
$Document.BuiltinDocumentProperties.Subject = "This is subject aka subtitle"

$null = $Document.AddHyperLink("Google", [uri] "https://www.google.com")

$Null = $Document.AddHyperLink("Evotec", [uri] "https://evotec.xyz", $true, "Tooltip for hyperlink", $false)


$Document.HyperLinks.Count

foreach ($HyperLink in $Document.HyperLinks) {
    if ($HyperLink.IsEmail) {
        Write-Host -Object "Email: $($HyperLink.EmailAddress)"
    } elseif ($HyperLink.IsHttp) {
        Write-Host -Object "URL: $($HyperLink.Uri)"
    } else {
        Write-Host -Object "Text: $($HyperLink.Text)"
    }
    # display properties of hyerlink
    $HyperLink | Format-Table

}

Save-OfficeWord -Document $Document -Show