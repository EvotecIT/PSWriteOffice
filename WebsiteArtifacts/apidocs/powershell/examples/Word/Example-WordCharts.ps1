Import-Module (Join-Path $PSScriptRoot '..\..\PSWriteOffice.psd1') -Force

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$docPath = Join-Path $documents 'Word-Charts.docx'
$doc = New-OfficeWord -Path $docPath

try {
    $doc.AddParagraph('Word charts are currently available through the underlying OfficeIMO document object.') | Out-Null
    $doc.AddParagraph('Older samples that used PieChart() should now use AddChart().AddPie().') | Out-Null

    $chart = $doc.AddChart('Regional Revenue Mix')
    $chart.AddPie('North America', 125000).
        AddPie('EMEA', 98000).
        AddPie('APAC', 143000) | Out-Null
    $chart.SetWidthToPageContent(0.70, 320) | Out-Null

    Close-OfficeWord -Document $doc -Save
    $doc = $null
} finally {
    if ($null -ne $doc) {
        $doc.Dispose()
    }
}

Write-Host "Document saved to $docPath"
