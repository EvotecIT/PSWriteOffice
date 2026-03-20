---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeMarkdownHtml
## SYNOPSIS
Converts Markdown content to HTML.

## SYNTAX
### Path (Default)
```powershell
ConvertTo-OfficeMarkdownHtml [-InputPath] <string> [-OutputPath <string>] [-DocumentMode] [-Style <HtmlStyle>] [-CssDelivery <CssDelivery>] [-AssetMode <AssetMode>] [-Title <string>] [-ReaderOptions <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-PassThru] [<CommonParameters>]
```

### Text
```powershell
ConvertTo-OfficeMarkdownHtml -Text <string> [-OutputPath <string>] [-DocumentMode] [-Style <HtmlStyle>] [-CssDelivery <CssDelivery>] [-AssetMode <AssetMode>] [-Title <string>] [-ReaderOptions <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
ConvertTo-OfficeMarkdownHtml -Document <MarkdownDoc> [-OutputPath <string>] [-DocumentMode] [-Style <HtmlStyle>] [-CssDelivery <CssDelivery>] [-AssetMode <AssetMode>] [-Title <string>] [-ReaderOptions <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Converts Markdown content to HTML.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$html = ConvertTo-OfficeMarkdownHtml -Path .\README.md
```

Returns the rendered HTML.

### EXAMPLE 2
```powershell
PS>ConvertTo-OfficeMarkdownHtml -Path .\Report.md -DocumentMode -Title 'Weekly Report' -Style Clean -OutputPath .\Report.html -PassThru
```

Generates a full HTML file with title and CSS styling.

## PARAMETERS

### -AssetMode
Asset loading mode.

```yaml
Type: AssetMode
Parameter Sets: Path, Text, Document
Aliases: None
Possible values: Online, Offline

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CssDelivery
CSS delivery mode.

```yaml
Type: CssDelivery
Parameter Sets: Path, Text, Document
Aliases: None
Possible values: Inline, ExternalFile, LinkHref, None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Markdown document to convert.

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

### -DocumentMode
Render a full HTML document instead of a fragment.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
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

### -OutputPath
Optional output path for the HTML file.

```yaml
Type: String
Parameter Sets: Path, Text, Document
Aliases: OutPath
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit a FileInfo when saving to disk.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Profile
Named reader profile used when ReaderOptions is not supplied.

```yaml
Type: Nullable`1
Parameter Sets: Path, Text, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ReaderOptions
Optional reader options when parsing Markdown.

```yaml
Type: MarkdownReaderOptions
Parameter Sets: Path, Text, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Style
Built-in HTML style preset.

```yaml
Type: HtmlStyle
Parameter Sets: Path, Text, Document
Aliases: None
Possible values: Plain, Clean, GithubLight, GithubDark, GithubAuto, ChatLight, ChatDark, ChatAuto, Word

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
Markdown text to convert.

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

### -Title
Optional title for HTML documents.

```yaml
Type: String
Parameter Sets: Path, Text, Document
Aliases: None
Possible values: 

Required: False
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

- `System.String
System.IO.FileInfo`

## RELATED LINKS

- None

