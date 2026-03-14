---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeWordHtml
## SYNOPSIS
Converts a Word document to HTML.

## SYNTAX
### Path (Default)
```powershell
ConvertTo-OfficeWordHtml [-FilePath] <string> [-OutputPath <string>] [-FontFamily <string>] [-IncludeFontStyles] [-IncludeListStyles] [-IncludeParagraphClasses] [-IncludeRunClasses] [-IncludeDefaultCss] [-UseImagePaths] [-ExcludeFootnotes] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
ConvertTo-OfficeWordHtml -Document <WordDocument> [-OutputPath <string>] [-FontFamily <string>] [-IncludeFontStyles] [-IncludeListStyles] [-IncludeParagraphClasses] [-IncludeRunClasses] [-IncludeDefaultCss] [-UseImagePaths] [-ExcludeFootnotes] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Converts a Word document to HTML.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$html = ConvertTo-OfficeWordHtml -Path .\report.docx
```

Loads the document and returns HTML markup.

### EXAMPLE 2
```powershell
PS>ConvertTo-OfficeWordHtml -Path .\report.docx -OutputPath .\report.html -PassThru
```

Writes report.html and returns the file info.

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

### -ExcludeFootnotes
Exclude footnotes from the HTML output.

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
Optional font family to use during conversion.

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

### -IncludeDefaultCss
Include the built-in default CSS in the HTML head.

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

### -IncludeFontStyles
Include font styles as inline CSS.

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

### -IncludeListStyles
Include list style metadata.

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

### -IncludeParagraphClasses
Emit paragraph styles as CSS classes.

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

### -IncludeRunClasses
Emit run styles as CSS classes.

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

### -OutputPath
Optional output path for the HTML file.

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

### -UseImagePaths
Store image references as file paths instead of base64 data URIs.

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

