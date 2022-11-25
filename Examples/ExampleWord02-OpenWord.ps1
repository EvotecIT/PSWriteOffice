Clear-Host
Import-Module .\PSWriteOffice.psd1 -Force

if (Test-Path -LiteralPath "$PSScriptRoot\Documents\Test5.docx") {
    $Document = Get-OfficeWord -FilePath $PSScriptRoot\Documents\Test5.docx
    $Document.BuiltinDocumentProperties.Category = 'Test'
    $Document.ApplicationProperties.Company = "Evotec"

    New-OfficeWordText -Document $Document -Text 'Add more things!' -Bold $null, $true -Underline $null, 'Dotted'
    New-OfficeWordText -Document $Document -Text 'Add more things!', ' Ok?' -Bold $true -Underline $null, 'Double'
    New-OfficeWordText -Document $Document -Text 'Add more things!', ' a bit more  with bold' -Bold $null, $true -Underline Dash -Color Red

    #$Document.Settings.UpdateFieldsOnOpen = $true
    Save-OfficeWord -Document $Document -Show -FilePath $PSScriptRoot\Documents\Test6.docx -Retry 1
}