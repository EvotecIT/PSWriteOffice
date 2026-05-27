---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeMarkdownHtml
## SYNOPSIS
Converts HTML content to Markdown.

## SYNTAX
### Html (Default)
```powershell
ConvertFrom-OfficeMarkdownHtml [-Html] <string> [-OutputPath <string>] [-AsDocument] [-PassThru] [-Options <HtmlToMarkdownOptions>] [-Portable] [-BaseUri <string>] [-IncludeDocumentChrome] [-PreserveScriptsAndStyles] [-DropUnsupportedBlocks] [-DropUnsupportedInlineHtml] [-MaxInputCharacters <int>] [<CommonParameters>]
```

### Path
```powershell
ConvertFrom-OfficeMarkdownHtml [-InputPath] <string> [-OutputPath <string>] [-AsDocument] [-PassThru] [-Options <HtmlToMarkdownOptions>] [-Portable] [-BaseUri <string>] [-IncludeDocumentChrome] [-PreserveScriptsAndStyles] [-DropUnsupportedBlocks] [-DropUnsupportedInlineHtml] [-MaxInputCharacters <int>] [<CommonParameters>]
```

## DESCRIPTION
Returns Markdown text by default. Use `-OutputPath` to save Markdown to disk, or `-AsDocument` to return an `OfficeIMO.Markdown.MarkdownDoc` for further editing and rendering.

This command is for static HTML fragments and documents. For browser-backed scraping, table extraction, metadata extraction, or rendered-page workflows, use PSParseHTML first and pass the resulting objects or Markdown into PSWriteOffice.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $markdown = ConvertFrom-OfficeMarkdownHtml -Html '<h1>Report</h1><p>Ready</p>'
```

Returns Markdown converted from the supplied HTML fragment.

### EXAMPLE 2
```powershell
PS> ConvertFrom-OfficeMarkdownHtml -Path .\report.html -Portable -OutputPath .\report.md -PassThru
```

Saves portable Markdown converted from an HTML file and returns the created file.

### EXAMPLE 3
```powershell
PS> $doc = ConvertFrom-OfficeMarkdownHtml -Path .\report.html -AsDocument
```

Returns a Markdown document object for additional edits or conversion to HTML/Word.

## PARAMETERS

### -AsDocument
Emit a Markdown document object instead of Markdown text.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Required: False
Position: named
Accept pipeline input: False
```

### -BaseUri
Base URI used to resolve relative links and image sources.

```yaml
Type: String
Parameter Sets: Html, Path
Required: False
Position: named
Accept pipeline input: False
```

### -DropUnsupportedBlocks
Drop unsupported block HTML instead of preserving it as raw HTML.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Required: False
Position: named
Accept pipeline input: False
```

### -DropUnsupportedInlineHtml
Drop unsupported inline HTML instead of preserving it as raw HTML.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Required: False
Position: named
Accept pipeline input: False
```

### -Html
HTML markup to convert.

```yaml
Type: String
Parameter Sets: Html
Required: True
Position: 0
Accept pipeline input: True (ByValue)
```

### -IncludeDocumentChrome
Convert the full HTML document instead of only body contents.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Required: False
Position: named
Accept pipeline input: False
```

### -InputPath
Path to an HTML file.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath, Path
Required: True
Position: 0
Accept pipeline input: False
```

### -MaxInputCharacters
Maximum input length, in characters, accepted by the converter.

```yaml
Type: Nullable`1
Parameter Sets: Html, Path
Required: False
Position: named
Accept pipeline input: False
```

### -Options
Optional conversion options.

```yaml
Type: HtmlToMarkdownOptions
Parameter Sets: Html, Path
Required: False
Position: named
Accept pipeline input: False
```

### -OutputPath
Optional output path for the Markdown file.

```yaml
Type: String
Parameter Sets: Html, Path
Aliases: OutPath
Required: False
Position: named
Accept pipeline input: False
```

### -PassThru
Emit a FileInfo when saving to disk.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Required: False
Position: named
Accept pipeline input: False
```

### -Portable
Use portable Markdown output when `-Options` is not supplied.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Required: False
Position: named
Accept pipeline input: False
```

### -PreserveScriptsAndStyles
Preserve script, style, noscript, and template elements.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Required: False
Position: named
Accept pipeline input: False
```

## RELATED LINKS

[ConvertTo-OfficeMarkdownHtml](ConvertTo-OfficeMarkdownHtml.md)
[ConvertFrom-OfficeWordHtml](ConvertFrom-OfficeWordHtml.md)
