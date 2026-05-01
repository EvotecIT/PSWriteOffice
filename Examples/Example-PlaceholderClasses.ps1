$modulePath = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
    $env:PSWRITEOFFICE_MODULE_MANIFEST
} else {
    "$PSScriptRoot/../PSWriteOffice.psd1"
}
if (-not (Get-Module -Name PSWriteOffice)) { Import-Module $modulePath -ErrorAction Stop }
# Instantiate placeholder service and cmdlet
$service = [PSWriteOffice.Services.Word.WordDocumentService]::new()
$cmdlet = [PSWriteOffice.Cmdlets.Word.GetOfficeWordCommand]::new()

$service
$cmdlet
