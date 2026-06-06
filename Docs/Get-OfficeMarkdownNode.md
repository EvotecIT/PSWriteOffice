---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeMarkdownNode
## SYNOPSIS
Gets the OfficeIMO.Markdown object tree from Markdown content.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeMarkdownNode [-InputPath] <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-NodeType <string>] [-MaxDepth <int>] [-CaseSensitive] [-Raw] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeMarkdownNode -Document <MarkdownDoc> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-NodeType <string>] [-MaxDepth <int>] [-CaseSensitive] [-Raw] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeMarkdownNode -Text <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-NodeType <string>] [-MaxDepth <int>] [-CaseSensitive] [-Raw] [<CommonParameters>]
```

## DESCRIPTION
Returns PowerShell-friendly node records by default. Use -Raw to emit the underlying OfficeIMO nodes.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeMarkdownNode -Text "# Report`n`n## Summary"
```

Parses Markdown text and returns the document, block, and inline object tree.

### EXAMPLE 2
```powershell
PS> Get-OfficeMarkdown -Path .\README.md | Get-OfficeMarkdownNode -NodeType '*Table*'
```

Returns matching nodes from an existing Markdown document.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for node type filters.

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

### -MaxDepth
Maximum traversal depth. Zero returns only the document root.

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

### -NodeType
Optional wildcard pattern matched against node type names.

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

### -Raw
Emit raw OfficeIMO.Markdown node objects instead of PowerShell-friendly records.

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

- `System.Management.Automation.PSObject
OfficeIMO.Markdown.MarkdownObject`

## RELATED LINKS

- None
