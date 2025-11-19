Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
$docPath = Join-Path $documents 'Word-AliasDsl.docx'

if (-not (Test-Path $docPath)) {
    Write-Warning "Expected document '$docPath' not found. Run Example-WordAliasDsl.ps1 first."
    return
}

$document = Get-OfficeWord -Path $docPath -ReadOnly
try {
    Write-Host "Sections:" $document.Sections.Count
    Write-Host 'First paragraphs:'
    $document.Paragraphs | Select-Object -First 3 | ForEach-Object {
        Write-Host " -" $_.Text
    }
} finally {
    Close-OfficeWord -Document $document
}
