---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeMarkdownFrontMatter
## SYNOPSIS
Gets YAML front matter entries from a Markdown document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeMarkdownFrontMatter [-InputPath] <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-Key <string>] [-CaseSensitive] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeMarkdownFrontMatter -Document <MarkdownDoc> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-Key <string>] [-CaseSensitive] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeMarkdownFrontMatter -Text <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownReaderOptions+MarkdownDialectProfile>] [-Key <string>] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Gets YAML front matter entries from a Markdown document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $metadata = Get-OfficeMarkdownFrontMatter -Path .\Report.md -Key 'title'
            $metadata |
                Select-Object -Property Key, Value |
                Format-Table -AutoSize
```

Parses a Markdown file and returns matching front matter entries for metadata proof.

### EXAMPLE 2
```powershell
PS> $markdown = "---`ntitle: Report`nstatus: Ready`n---`n# Report"
            Get-OfficeMarkdownFrontMatter -Text $markdown |
                Select-Object -Property Key, Value
```

Parses Markdown text directly when the document has not been saved yet.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for key filters.

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

### -Key
Optional wildcard pattern matched against front matter keys.

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

- `OfficeIMO.Markdown.FrontMatterBlock+Entry`

## RELATED LINKS

- None
