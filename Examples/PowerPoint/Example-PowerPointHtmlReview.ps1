Import-Module PSWriteOffice -ErrorAction Stop

$documents = Join-Path $PSScriptRoot '..\Documents'
New-Item -Path $documents -ItemType Directory -Force | Out-Null

$presentationPath = Join-Path $documents 'PowerPoint-HtmlReview.pptx'
$semanticHtmlPath = Join-Path $documents 'PowerPoint-HtmlReview.semantic.html'
$visualHtmlPath = Join-Path $documents 'PowerPoint-HtmlReview.visual.html'

$presentation = New-OfficePowerPoint -FilePath $presentationPath

$statusSlide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Set-OfficePowerPointSlideTitle -Slide $statusSlide -Title 'Monthly Service Review' | Out-Null
Add-OfficePowerPointTextBox -Slide $statusSlide -Text 'Identity, Messaging, and Reporting are ready for leadership review.' -X 80 -Y 140 -Width 560 -Height 80 | Out-Null
Set-OfficePowerPointNotes -Slide $statusSlide -Text 'Use this slide to introduce the operational status summary.' | Out-Null

$tableSlide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Set-OfficePowerPointSlideTitle -Slide $tableSlide -Title 'Open Items' | Out-Null
Add-OfficePowerPointTable -Slide $tableSlide -Headers 'Area', 'Owner', 'Next Step' -Rows @(
    @('Messaging', 'Collaboration', 'Review retry spikes')
    @('Reporting', 'Analytics', 'Publish refreshed dashboard')
) -X 70 -Y 130 -Width 600 -Height 160 | Out-Null

Save-OfficePowerPoint -Presentation $presentation
$presentation.Dispose()

ConvertTo-OfficePowerPointHtml -Path $presentationPath -OutputPath $semanticHtmlPath -Title 'Deck Review' -PassThru | Out-Null
ConvertTo-OfficePowerPointHtml -Path $presentationPath -Profile VisualReview -OutputPath $visualHtmlPath -Title 'Deck Visual Review' -PassThru | Out-Null

Write-Host "Presentation saved to $presentationPath"
Write-Host "Semantic HTML saved to $semanticHtmlPath"
Write-Host "Visual HTML saved to $visualHtmlPath"
