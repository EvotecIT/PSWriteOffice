Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

# check if document exists
if (Test-Path -LiteralPath "$PSScriptRoot\Documents\BasicDocument.docx") {
    # load document
    $Document = Get-OfficeWord -FilePath $PSScriptRoot\Documents\BasicDocument.docx

    # extract text from all paragraphs at once
    $Document.Paragraphs.Text

    # if tables exists you can extract data from them as well
    $Document.Tables[0]
}

Close-OfficeWord -Document $Document