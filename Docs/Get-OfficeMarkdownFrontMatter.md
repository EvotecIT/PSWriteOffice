# Get-OfficeMarkdownFrontMatter

Gets YAML front matter entries from a Markdown document.

## Synopsis

```powershell
Get-OfficeMarkdownFrontMatter [-InputPath] <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-Key <string>] [-CaseSensitive]
Get-OfficeMarkdownFrontMatter -Text <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-Key <string>] [-CaseSensitive]
Get-OfficeMarkdownFrontMatter -Document <MarkdownDoc> [-Key <string>] [-CaseSensitive]
```

## Description

Parses Markdown or accepts an `OfficeIMO.Markdown.MarkdownDoc` from the pipeline and returns structured front matter entries with `Key` and `Value`.

## Examples

```powershell
Get-OfficeMarkdownFrontMatter -Path .\post.md
```

```powershell
Get-OfficeMarkdown -Path .\post.md | Get-OfficeMarkdownFrontMatter -Key title
```

## Outputs

- `OfficeIMO.Markdown.FrontMatterBlock+Entry`
