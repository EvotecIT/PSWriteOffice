# Get-OfficeMarkdownHeading

Gets heading metadata from a Markdown document.

## Synopsis

```powershell
Get-OfficeMarkdownHeading [-InputPath] <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive]
Get-OfficeMarkdownHeading -Text <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive]
Get-OfficeMarkdownHeading -Document <MarkdownDoc> [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive]
```

## Description

Parses Markdown or accepts an `OfficeIMO.Markdown.MarkdownDoc` from the pipeline and returns heading metadata with `Level`, `Text`, `Anchor`, and `Block`.

## Examples

```powershell
Get-OfficeMarkdownHeading -Path .\README.md -MinLevel 2
```

```powershell
Get-OfficeMarkdown -Path .\README.md | Get-OfficeMarkdownHeading -HeadingText '*Install*'
```

## Outputs

- `OfficeIMO.Markdown.MarkdownDoc+HeadingInfo`
