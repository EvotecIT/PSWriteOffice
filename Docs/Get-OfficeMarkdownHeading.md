---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeMarkdownHeading
## SYNOPSIS
Gets heading metadata from a Markdown document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeMarkdownHeading [-InputPath] <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeMarkdownHeading -Document <MarkdownDoc> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeMarkdownHeading -Text <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Returns heading level, text, resolved anchor, and the backing heading block.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeMarkdownHeading -Text "# Title`n`n## Details"
```

Parses Markdown text and returns the document headings.

### EXAMPLE 2
```powershell
PS> Get-OfficeMarkdown -Path .\README.md | Get-OfficeMarkdownHeading -MinLevel 2
```

Returns headings from an existing Markdown document object.

## PARAMETERS

### -Anchor
Optional wildcard pattern matched against resolved heading anchors.

```yaml
Type: String
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CaseSensitive
Use case-sensitive matching for text and anchor filters.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Markdown document to inspect.

```yaml
Type: MarkdownDoc
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -HeadingText
Optional wildcard pattern matched against heading text.

```yaml
Type: String
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputPath
Path to the Markdown file.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath, Path
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxLevel
Maximum heading level to return.

```yaml
Type: Int32
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MinLevel
Minimum heading level to return.

```yaml
Type: Int32
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Options
Optional reader options used when parsing path or text input.

```yaml
Type: MarkdownReaderOptions
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Profile
Named reader profile used when Options is not supplied.

```yaml
Type: Nullable`1
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
Markdown text to parse.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc+HeadingInfo`

## RELATED LINKS

- None
