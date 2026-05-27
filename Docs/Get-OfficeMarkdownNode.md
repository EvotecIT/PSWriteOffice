# Get-OfficeMarkdownNode

Gets the OfficeIMO.Markdown object tree from Markdown content.

## Synopsis

```powershell
Get-OfficeMarkdownNode [-InputPath] <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-NodeType <string>] [-MaxDepth <int>] [-CaseSensitive] [-Raw]
Get-OfficeMarkdownNode -Text <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-NodeType <string>] [-MaxDepth <int>] [-CaseSensitive] [-Raw]
Get-OfficeMarkdownNode -Document <MarkdownDoc> [-NodeType <string>] [-MaxDepth <int>] [-CaseSensitive] [-Raw]
```

## Description

Parses Markdown or accepts an `OfficeIMO.Markdown.MarkdownDoc` from the pipeline and walks the OfficeIMO object tree: document, front matter, blocks, lists, list items, table cells, inline sequences, and supported inline nodes.

By default it emits PowerShell-friendly records with `Depth`, `Path`, `Type`, `SourceSpan`, `Text`, `Markdown`, and the underlying `Node`. Use `-Raw` to emit the raw OfficeIMO nodes.

## Examples

```powershell
Get-OfficeMarkdownNode -Path .\README.md
```

```powershell
Get-OfficeMarkdownNode -Path .\README.md -NodeType '*Table*'
```

```powershell
Get-OfficeMarkdown -Path .\README.md | Get-OfficeMarkdownNode -MaxDepth 2
```

## Outputs

- `System.Management.Automation.PSObject`
- `OfficeIMO.Markdown.MarkdownObject`
