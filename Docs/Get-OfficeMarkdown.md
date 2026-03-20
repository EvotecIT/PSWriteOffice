---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeMarkdown
## SYNOPSIS
Parses Markdown text or files into a Markdown document model.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeMarkdown [-InputPath] <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeMarkdown -Text <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [<CommonParameters>]
```

## DESCRIPTION
Parses Markdown text or files into a Markdown document model.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$md = Get-OfficeMarkdown -Path .\README.md
```

Loads the file into a Markdown document object.

### EXAMPLE 2
```powershell
PS>$md = Get-OfficeMarkdown -Text '# Title`n`nBody text'
```

Parses Markdown text directly into a document model.

## PARAMETERS

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
Optional reader options.

```yaml
Type: MarkdownReaderOptions
Parameter Sets: Path, Text
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
Parameter Sets: Path, Text
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

- `None`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None

