---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeWordMarkdown
## SYNOPSIS
Converts a Word document to Markdown.

## SYNTAX
### Path (Default)
```powershell
ConvertTo-OfficeWordMarkdown [-FilePath] <string> [-OutputPath <string>] [-FontFamily <string>] [-EnableUnderline] [-EnableHighlight] [-ImageExportMode <ImageExportMode>] [-ImageDirectory <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
ConvertTo-OfficeWordMarkdown -Document <WordDocument> [-OutputPath <string>] [-FontFamily <string>] [-EnableUnderline] [-EnableHighlight] [-ImageExportMode <ImageExportMode>] [-ImageDirectory <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Converts a Word document to Markdown.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$markdown = ConvertTo-OfficeWordMarkdown -Path .\report.docx
```

Loads the document and returns Markdown markup.

### EXAMPLE 2
```powershell
PS>ConvertTo-OfficeWordMarkdown -Path .\report.docx -OutputPath .\report.md -PassThru
```

Writes report.md and returns the file info.

## PARAMETERS

### -Document
Word document instance to convert.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -EnableHighlight
Wrap highlighted text with Markdown highlight markers.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -EnableUnderline
Wrap underlined text with HTML underline tags.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FilePath
Path to a .docx file.

```yaml
Type: String
Parameter Sets: Path
Aliases: Path
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FontFamily
Optional font family that should be treated as inline code.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ImageDirectory
Directory used when exporting images as files.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ImageExportMode
Controls how images are emitted during Markdown conversion.

```yaml
Type: ImageExportMode
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: Base64
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Optional output path for the Markdown file.

```yaml
Type: String
Parameter Sets: Path, Document
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
Parameter Sets: Path, Document
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

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `System.String
System.IO.FileInfo`

## RELATED LINKS

- None
