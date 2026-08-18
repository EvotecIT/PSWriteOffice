---
title: "Read and Inspect Word Documents"
description: "Open existing DOCX files, enumerate document structures, search content, and collect statistics without changing the source document."
layout: docs
---

Use the Word read surface when the document already exists and the job is to understand it: inventory paragraphs and tables, locate terms, inspect links or controls, or decide whether a later update is safe.

## Open read-only

`Get-OfficeWord -ReadOnly` returns the OfficeIMO document model. Dispose it when finished. The targeted `Get-OfficeWord*` commands accept that object, so the file is opened once even when several structures are inspected.

```powershell
$document = Get-OfficeWord -Path '.\Input\Policy.docx' -ReadOnly
try {
    $paragraphs = @($document | Get-OfficeWordParagraph)
    $tables = @($document | Get-OfficeWordTable)
    $links = @($document | Get-OfficeWordHyperlink)
    $statistics = Get-OfficeWordStatistics -Document $document
} finally {
    Close-OfficeWord -Document $document
}
```

## Find content deliberately

Use `Find-OfficeWord -Text` for literal text and `-Pattern` for regular expressions. The inspection family also covers fields, bookmarks, comments and revisions, footnotes and endnotes, images, content controls, tables of contents, and document properties.

The complete [inspect-existing recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Word/Recipe-Word-InspectExisting.ps1) creates a representative file, opens it once, and reports sections, paragraphs, tables, word count, and matches.

## Choose the next step

- Use [targeted updates](/docs/pswriteoffice/word-update-existing/) when the structure should stay intact.
- Use [merge and mail merge](/docs/pswriteoffice/word-merge-mailmerge/) when several documents or data records produce the result.
- Use the cross-format [Reader](/docs/pswriteoffice/reader/) when the same search must include Word, Excel, PowerPoint, PDF, Markdown, email, and other supported files.
