# Get-OfficeMarkdownTable

Gets Markdown tables from a Markdown document.

## Synopsis

```powershell
Get-OfficeMarkdownTable [-InputPath] <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-AsObject]
Get-OfficeMarkdownTable -Text <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-AsObject]
Get-OfficeMarkdownTable -Document <MarkdownDoc> [-AsObject]
```

## Description

Parses Markdown or accepts an `OfficeIMO.Markdown.MarkdownDoc` from the pipeline and returns table blocks. Use `-AsObject` to emit table rows as PowerShell objects using Markdown table headers as property names.

## Examples

```powershell
Get-OfficeMarkdownTable -Path .\report.md
```

```powershell
Get-OfficeMarkdown -Path .\report.md | Get-OfficeMarkdownTable -AsObject
```

## Outputs

- `OfficeIMO.Markdown.TableBlock`
- `System.Management.Automation.PSObject`
