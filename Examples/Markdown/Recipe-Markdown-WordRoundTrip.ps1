$markdownPath = '.\Markdown-Word-Source.md'
$wordPath = '.\Markdown-Converted.docx'
$roundTripPath = '.\Markdown-Word-RoundTrip.md'
MarkdownNew -Path $markdownPath {
    MarkdownHeading -Level 1 -Text 'Change plan'
    MarkdownParagraph -Text 'The same source can be reviewed in Word and returned to Markdown.'
    MarkdownList -Items 'Prepare', 'Approve', 'Deploy'
}

ConvertFrom-OfficeWordMarkdown -FilePath $markdownPath -OutputPath $wordPath
ConvertTo-OfficeWordMarkdown -FilePath $wordPath -OutputPath $roundTripPath
