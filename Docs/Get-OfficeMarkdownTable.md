---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeMarkdownTable
## SYNOPSIS
Gets Markdown tables from a Markdown document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeMarkdownTable [-InputPath] <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-AsObject] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeMarkdownTable -Document <MarkdownDoc> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-AsObject] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeMarkdownTable -Text <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-AsObject] [<CommonParameters>]
```

## DESCRIPTION
Gets Markdown tables from a Markdown document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeMarkdown -Path .\Report.md | Get-OfficeMarkdownTable -AsObject
```

Returns table rows as PowerShell objects using the Markdown header row as property names.

## PARAMETERS

### -AsObject
Emit table rows as PowerShell objects instead of raw table blocks.

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

- `OfficeIMO.Markdown.TableBlock
System.Management.Automation.PSObject`

## RELATED LINKS

- None
