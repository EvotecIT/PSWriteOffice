Import-Module "$PSScriptRoot/../PSWriteOffice.psd1" -Force

# Instantiate placeholder service and cmdlet
$service = [PSWriteOffice.Services.Word.WordDocumentService]::new()
$cmdlet = [PSWriteOffice.Cmdlets.Word.GetOfficeWordCommand]::new()

$service
$cmdlet
